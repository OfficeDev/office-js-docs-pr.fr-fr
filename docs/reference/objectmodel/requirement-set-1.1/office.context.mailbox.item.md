
# <a name="item"></a><span data-ttu-id="9b9e6-101">item</span><span class="sxs-lookup"><span data-stu-id="9b9e6-101">item</span></span>

### <span data-ttu-id="9b9e6-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="9b9e6-p102">L’espace de noms `item` est utilisé pour accéder à vos messages, à vos demandes de réunion ou à vos rendez-vous. Vous pouvez déterminer le type de `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-106">Requirements</span></span>

|<span data-ttu-id="9b9e6-107">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-107">Requirement</span></span>| <span data-ttu-id="9b9e6-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-109">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-110">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-110">1.0</span></span>|
|[<span data-ttu-id="9b9e6-111">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-112">Restreint</span><span class="sxs-lookup"><span data-stu-id="9b9e6-112">Restricted</span></span>|
|[<span data-ttu-id="9b9e6-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="9b9e6-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-115">Example</span></span>

<span data-ttu-id="9b9e6-116">Cet exemple de code JavaScript montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9b9e6-117">Membres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="9b9e6-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9b9e6-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="9b9e6-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-121">Certains types de fichiers sont bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas rendus.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9b9e6-122">Pour plus d’information, voir les [pièces jointes bloquées par Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-123">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-123">Type:</span></span>

*   <span data-ttu-id="9b9e6-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9b9e6-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-125">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-125">Requirements</span></span>

|<span data-ttu-id="9b9e6-126">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-126">Requirement</span></span>| <span data-ttu-id="9b9e6-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-128">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-129">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-129">1.0</span></span>|
|[<span data-ttu-id="9b9e6-130">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-131">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-134">Example</span></span>

<span data-ttu-id="9b9e6-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9b9e6-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9b9e6-137">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9b9e6-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-139">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-139">Type:</span></span>

*   [<span data-ttu-id="9b9e6-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="9b9e6-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9b9e6-141">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-141">Requirements</span></span>

|<span data-ttu-id="9b9e6-142">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-142">Requirement</span></span>| <span data-ttu-id="9b9e6-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-144">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-145">1.1</span><span class="sxs-lookup"><span data-stu-id="9b9e6-145">1.1</span></span>|
|[<span data-ttu-id="9b9e6-146">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-147">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-149">Composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="9b9e6-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="9b9e6-152">Obtient un objet qui fournit des méthodes permettant de manipuler le texte d’un élément.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-153">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-153">Type:</span></span>

*   [<span data-ttu-id="9b9e6-154">Body</span><span class="sxs-lookup"><span data-stu-id="9b9e6-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="9b9e6-155">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-155">Requirements</span></span>

|<span data-ttu-id="9b9e6-156">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-156">Requirement</span></span>| <span data-ttu-id="9b9e6-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-158">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9b9e6-159">1.1</span></span>|
|[<span data-ttu-id="9b9e6-160">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-161">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9b9e6-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9b9e6-165">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9b9e6-166">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-167">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-167">Read mode</span></span>

<span data-ttu-id="9b9e6-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-170">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-170">Compose mode</span></span>

<span data-ttu-id="9b9e6-171">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-172">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-172">Type:</span></span>

*   <span data-ttu-id="9b9e6-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-174">Requirements</span></span>

|<span data-ttu-id="9b9e6-175">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-175">Requirement</span></span>| <span data-ttu-id="9b9e6-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-177">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-178">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-178">1.0</span></span>|
|[<span data-ttu-id="9b9e6-179">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-180">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-181">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-182">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-183">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="9b9e6-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="9b9e6-185">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9b9e6-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’identificateur de conversation de ce message changera et la valeur que vous aurez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9b9e6-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renverra une valeur.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-190">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-190">Type:</span></span>

*   <span data-ttu-id="9b9e6-191">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-192">Requirements</span></span>

|<span data-ttu-id="9b9e6-193">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-193">Requirement</span></span>| <span data-ttu-id="9b9e6-194">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-195">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-196">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-196">1.0</span></span>|
|[<span data-ttu-id="9b9e6-197">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-198">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="9b9e6-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="9b9e6-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="9b9e6-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-204">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-204">Type:</span></span>

*   <span data-ttu-id="9b9e6-205">Date</span><span class="sxs-lookup"><span data-stu-id="9b9e6-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-206">Requirements</span></span>

|<span data-ttu-id="9b9e6-207">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-207">Requirement</span></span>| <span data-ttu-id="9b9e6-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-209">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-210">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-210">1.0</span></span>|
|[<span data-ttu-id="9b9e6-211">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-212">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-214">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9b9e6-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="9b9e6-216">dateTimeModified :Date</span></span>

<span data-ttu-id="9b9e6-p111">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-219">Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-220">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-220">Type:</span></span>

*   <span data-ttu-id="9b9e6-221">Date</span><span class="sxs-lookup"><span data-stu-id="9b9e6-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-222">Requirements</span></span>

|<span data-ttu-id="9b9e6-223">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-223">Requirement</span></span>| <span data-ttu-id="9b9e6-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-225">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-226">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-226">1.0</span></span>|
|[<span data-ttu-id="9b9e6-227">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-228">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-230">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-231">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="9b9e6-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="9b9e6-233">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9b9e6-p112">La propriété `end` est exprimée en date et heure U.T.C. (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-236">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-236">Read mode</span></span>

<span data-ttu-id="9b9e6-237">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-238">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-238">Compose mode</span></span>

<span data-ttu-id="9b9e6-239">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9b9e6-240">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-241">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-241">Type:</span></span>

*   <span data-ttu-id="9b9e6-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-243">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-243">Requirements</span></span>

|<span data-ttu-id="9b9e6-244">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-244">Requirement</span></span>| <span data-ttu-id="9b9e6-245">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-246">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-247">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-247">1.0</span></span>|
|[<span data-ttu-id="9b9e6-248">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-249">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-250">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-251">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-252">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-252">Example</span></span>

<span data-ttu-id="9b9e6-253">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9b9e6-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9b9e6-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9b9e6-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété expéditeur représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-259">La propriété  `recipientType` de l'objet  `EmailAddressDetails` dans la propriété  `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-260">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-260">Type:</span></span>

*   [<span data-ttu-id="9b9e6-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9b9e6-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9b9e6-262">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-262">Requirements</span></span>

|<span data-ttu-id="9b9e6-263">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-263">Requirement</span></span>| <span data-ttu-id="9b9e6-264">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-265">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-266">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-266">1.0</span></span>|
|[<span data-ttu-id="9b9e6-267">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-268">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-269">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-270">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="9b9e6-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-271">internetMessageId :String</span></span>

<span data-ttu-id="9b9e6-p115">Obtient l’identificateur de message Internet d’un e-mail. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-274">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-274">Type:</span></span>

*   <span data-ttu-id="9b9e6-275">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-276">Requirements</span></span>

|<span data-ttu-id="9b9e6-277">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-277">Requirement</span></span>| <span data-ttu-id="9b9e6-278">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-279">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-280">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-280">1.0</span></span>|
|[<span data-ttu-id="9b9e6-281">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-282">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-284">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-285">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9b9e6-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-286">itemClass :String</span></span>

<span data-ttu-id="9b9e6-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9b9e6-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9b9e6-291">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-291">Type</span></span> | <span data-ttu-id="9b9e6-292">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-292">Description</span></span> | <span data-ttu-id="9b9e6-293">Classe d’élément</span><span class="sxs-lookup"><span data-stu-id="9b9e6-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9b9e6-294">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="9b9e6-294">Appointment items</span></span> | <span data-ttu-id="9b9e6-295">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="9b9e6-296">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="9b9e6-296">Message items</span></span> | <span data-ttu-id="9b9e6-297">Ces éléments incluent les e-mails dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9b9e6-298">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-299">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-299">Type:</span></span>

*   <span data-ttu-id="9b9e6-300">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-301">Requirements</span></span>

|<span data-ttu-id="9b9e6-302">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-302">Requirement</span></span>| <span data-ttu-id="9b9e6-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-304">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-305">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-305">1.0</span></span>|
|[<span data-ttu-id="9b9e6-306">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-307">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9b9e6-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-311">(nullable) itemId :String</span></span>

<span data-ttu-id="9b9e6-p118">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-314">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9b9e6-315">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9b9e6-316">Avant que d’appeler l’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de `Office.context.mailbox.convertToRestId`, qui est disponible à partir de l’ensemble de conditions requises 1.3.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="9b9e6-317">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-318">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-318">Type:</span></span>

*   <span data-ttu-id="9b9e6-319">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-320">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-320">Requirements</span></span>

|<span data-ttu-id="9b9e6-321">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-321">Requirement</span></span>| <span data-ttu-id="9b9e6-322">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-323">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-324">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-324">1.0</span></span>|
|[<span data-ttu-id="9b9e6-325">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-326">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-327">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-328">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-329">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-329">Example</span></span>

<span data-ttu-id="9b9e6-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="9b9e6-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9b9e6-333">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9b9e6-334">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-335">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-335">Type:</span></span>

*   [<span data-ttu-id="9b9e6-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9b9e6-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9b9e6-337">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-337">Requirements</span></span>

|<span data-ttu-id="9b9e6-338">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-338">Requirement</span></span>| <span data-ttu-id="9b9e6-339">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-340">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-340">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-341">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-341">1.0</span></span>|
|[<span data-ttu-id="9b9e6-342">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-343">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-344">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-345">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-346">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="9b9e6-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="9b9e6-348">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-349">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-349">Read mode</span></span>

<span data-ttu-id="9b9e6-350">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-351">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-351">Compose mode</span></span>

<span data-ttu-id="9b9e6-352">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-353">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-353">Type:</span></span>

*   <span data-ttu-id="9b9e6-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-355">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-355">Requirements</span></span>

|<span data-ttu-id="9b9e6-356">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-356">Requirement</span></span>| <span data-ttu-id="9b9e6-357">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-358">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-359">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-359">1.0</span></span>|
|[<span data-ttu-id="9b9e6-360">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-361">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-362">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-363">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-364">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9b9e6-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-365">normalizedSubject :String</span></span>

<span data-ttu-id="9b9e6-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9b9e6-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-370">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-370">Type:</span></span>

*   <span data-ttu-id="9b9e6-371">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-372">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-372">Requirements</span></span>

|<span data-ttu-id="9b9e6-373">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-373">Requirement</span></span>| <span data-ttu-id="9b9e6-374">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-375">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-376">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-376">1.0</span></span>|
|[<span data-ttu-id="9b9e6-377">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-378">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-379">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-380">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-381">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9b9e6-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9b9e6-383">Fournit l’accès aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9b9e6-384">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-385">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-385">Read mode</span></span>

<span data-ttu-id="9b9e6-386">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-387">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-387">Compose mode</span></span>

<span data-ttu-id="9b9e6-388">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d'obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-389">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-389">Type:</span></span>

*   <span data-ttu-id="9b9e6-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-391">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-391">Requirements</span></span>

|<span data-ttu-id="9b9e6-392">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-392">Requirement</span></span>| <span data-ttu-id="9b9e6-393">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-394">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-394">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-395">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-395">1.0</span></span>|
|[<span data-ttu-id="9b9e6-396">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-397">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-398">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-399">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-400">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9b9e6-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9b9e6-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-404">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-404">Type:</span></span>

*   [<span data-ttu-id="9b9e6-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9b9e6-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9b9e6-406">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-406">Requirements</span></span>

|<span data-ttu-id="9b9e6-407">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-407">Requirement</span></span>| <span data-ttu-id="9b9e6-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-409">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-410">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-410">1.0</span></span>|
|[<span data-ttu-id="9b9e6-411">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-412">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-414">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9b9e6-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9b9e6-417">Fournit l’accès aux participants obligatoires d'un événement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9b9e6-418">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-419">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-419">Read mode</span></span>

<span data-ttu-id="9b9e6-420">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant obligatoires de la réunion.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-421">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-421">Compose mode</span></span>

<span data-ttu-id="9b9e6-422">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-423">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-423">Type:</span></span>

*   <span data-ttu-id="9b9e6-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-425">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-425">Requirements</span></span>

|<span data-ttu-id="9b9e6-426">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-426">Requirement</span></span>| <span data-ttu-id="9b9e6-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-428">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-429">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-429">1.0</span></span>|
|[<span data-ttu-id="9b9e6-430">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-431">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-433">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9b9e6-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9b9e6-p126">Obtient l’adresse de messagerie de l’expéditeur d’un e-mail. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9b9e6-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété expéditeur représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-440">La propriété  `recipientType` de l'objet  `EmailAddressDetails` dans la propriété  `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-441">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-441">Type:</span></span>

*   [<span data-ttu-id="9b9e6-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9b9e6-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9b9e6-443">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-443">Requirements</span></span>

|<span data-ttu-id="9b9e6-444">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-444">Requirement</span></span>| <span data-ttu-id="9b9e6-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-446">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-447">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-447">1.0</span></span>|
|[<span data-ttu-id="9b9e6-448">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-449">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-451">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-452">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="9b9e6-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="9b9e6-454">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9b9e6-p128">La propriété `start` est exprimée en date et heure U.T.C. (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-457">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-457">Read mode</span></span>

<span data-ttu-id="9b9e6-458">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-459">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-459">Compose mode</span></span>

<span data-ttu-id="9b9e6-460">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9b9e6-461">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format U.T.C. pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-462">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-462">Type:</span></span>

*   <span data-ttu-id="9b9e6-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-464">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-464">Requirements</span></span>

|<span data-ttu-id="9b9e6-465">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-465">Requirement</span></span>| <span data-ttu-id="9b9e6-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-467">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-468">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-468">1.0</span></span>|
|[<span data-ttu-id="9b9e6-469">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-470">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-472">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-473">Example</span></span>

<span data-ttu-id="9b9e6-474">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="9b9e6-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="9b9e6-476">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9b9e6-477">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-478">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-478">Read mode</span></span>

<span data-ttu-id="9b9e6-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-481">Compose mode</span></span>

<span data-ttu-id="9b9e6-482">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9b9e6-483">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-483">Type:</span></span>

*   <span data-ttu-id="9b9e6-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-485">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-485">Requirements</span></span>

|<span data-ttu-id="9b9e6-486">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-486">Requirement</span></span>| <span data-ttu-id="9b9e6-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-488">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-489">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-489">1.0</span></span>|
|[<span data-ttu-id="9b9e6-490">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-491">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-493">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9b9e6-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9b9e6-495">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9b9e6-496">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9b9e6-497">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-497">Read mode</span></span>

<span data-ttu-id="9b9e6-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9b9e6-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-500">Compose mode</span></span>

<span data-ttu-id="9b9e6-501">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9b9e6-502">Type :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-502">Type:</span></span>

*   <span data-ttu-id="9b9e6-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-504">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-504">Requirements</span></span>

|<span data-ttu-id="9b9e6-505">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-505">Requirement</span></span>| <span data-ttu-id="9b9e6-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-507">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-508">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-508">1.0</span></span>|
|[<span data-ttu-id="9b9e6-509">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-510">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-512">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-513">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="9b9e6-514">Méthodes</span><span class="sxs-lookup"><span data-stu-id="9b9e6-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9b9e6-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9b9e6-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9b9e6-516">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9b9e6-517">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9b9e6-518">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-519">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-519">Parameters:</span></span>

|<span data-ttu-id="9b9e6-520">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-520">Name</span></span>| <span data-ttu-id="9b9e6-521">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-521">Type</span></span>| <span data-ttu-id="9b9e6-522">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-522">Attributes</span></span>| <span data-ttu-id="9b9e6-523">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9b9e6-524">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-524">String</span></span>||<span data-ttu-id="9b9e6-p132">L’URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9b9e6-527">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-527">String</span></span>||<span data-ttu-id="9b9e6-p133">Nom de la pièce jointe affiché lors de son chargement. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9b9e6-530">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-530">Object</span></span>| <span data-ttu-id="9b9e6-531">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-531">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-532">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9b9e6-533">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-533">Object</span></span>| <span data-ttu-id="9b9e6-534">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-534">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-535">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9b9e6-536">function</span><span class="sxs-lookup"><span data-stu-id="9b9e6-536">function</span></span>| <span data-ttu-id="9b9e6-537">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-537">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-538">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9b9e6-539">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9b9e6-540">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9b9e6-541">Erreurs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-541">Errors</span></span>

| <span data-ttu-id="9b9e6-542">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-542">Error code</span></span> | <span data-ttu-id="9b9e6-543">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9b9e6-544">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9b9e6-545">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9b9e6-546">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9b9e6-547">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-547">Requirements</span></span>

|<span data-ttu-id="9b9e6-548">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-548">Requirement</span></span>| <span data-ttu-id="9b9e6-549">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-550">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-551">1.1</span><span class="sxs-lookup"><span data-stu-id="9b9e6-551">1.1</span></span>|
|[<span data-ttu-id="9b9e6-552">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="9b9e6-554">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-555">Composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-556">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9b9e6-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9b9e6-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9b9e6-558">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9b9e6-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9b9e6-562">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9b9e6-563">Si votre complément Office est exécuté dans la Outlook Web App , la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez, mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-564">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-564">Parameters:</span></span>

|<span data-ttu-id="9b9e6-565">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-565">Name</span></span>| <span data-ttu-id="9b9e6-566">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-566">Type</span></span>| <span data-ttu-id="9b9e6-567">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-567">Attributes</span></span>| <span data-ttu-id="9b9e6-568">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9b9e6-569">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-569">String</span></span>||<span data-ttu-id="9b9e6-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9b9e6-572">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-572">String</span></span>||<span data-ttu-id="9b9e6-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9b9e6-575">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-575">Object</span></span>| <span data-ttu-id="9b9e6-576">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-576">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-577">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9b9e6-578">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-578">Object</span></span>| <span data-ttu-id="9b9e6-579">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-579">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-580">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9b9e6-581">function</span><span class="sxs-lookup"><span data-stu-id="9b9e6-581">function</span></span>| <span data-ttu-id="9b9e6-582">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-582">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-583">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9b9e6-584">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9b9e6-585">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9b9e6-586">Erreurs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-586">Errors</span></span>

| <span data-ttu-id="9b9e6-587">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-587">Error code</span></span> | <span data-ttu-id="9b9e6-588">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9b9e6-589">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9b9e6-590">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-590">Requirements</span></span>

|<span data-ttu-id="9b9e6-591">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-591">Requirement</span></span>| <span data-ttu-id="9b9e6-592">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-593">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-594">1.1</span><span class="sxs-lookup"><span data-stu-id="9b9e6-594">1.1</span></span>|
|[<span data-ttu-id="9b9e6-595">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="9b9e6-597">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-598">Composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-599">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-599">Example</span></span>

<span data-ttu-id="9b9e6-600">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="9b9e6-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="9b9e6-602">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-603">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9b9e6-604">Sur Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9b9e6-605">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` lève une exception.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-606">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyAllForm` n’est pas prise en charge dans l’ensemble des conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-606">NOTE: The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="9b9e6-607">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-608">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-608">Parameters:</span></span>

|<span data-ttu-id="9b9e6-609">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-609">Name</span></span>| <span data-ttu-id="9b9e6-610">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-610">Type</span></span>| <span data-ttu-id="9b9e6-611">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9b9e6-612">String | Object</span><span class="sxs-lookup"><span data-stu-id="9b9e6-612">String &#124; Object</span></span>| |<span data-ttu-id="9b9e6-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9b9e6-615">**OU**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-615">**OR**</span></span><br/><span data-ttu-id="9b9e6-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9b9e6-618">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-618">String</span></span> | <span data-ttu-id="9b9e6-619">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-619">&lt;optional&gt;</span></span> | <span data-ttu-id="9b9e6-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="9b9e6-622">function</span><span class="sxs-lookup"><span data-stu-id="9b9e6-622">function</span></span> | <span data-ttu-id="9b9e6-623">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-623">&lt;optional&gt;</span></span> | <span data-ttu-id="9b9e6-624">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9b9e6-625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-625">Requirements</span></span>

|<span data-ttu-id="9b9e6-626">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-626">Requirement</span></span>| <span data-ttu-id="9b9e6-627">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-628">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-629">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-629">1.0</span></span>|
|[<span data-ttu-id="9b9e6-630">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-631">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-632">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-633">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9b9e6-634">Exemples</span><span class="sxs-lookup"><span data-stu-id="9b9e6-634">Examples</span></span>

<span data-ttu-id="9b9e6-635">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9b9e6-636">Réponse sans texte.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9b9e6-637">Réponse avec seulement un corps de message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9b9e6-638">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="9b9e6-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="9b9e6-640">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-641">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-641">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9b9e6-642">Sur Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9b9e6-643">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` lève une exception.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-644">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyForm` n’est pas prise en charge dans l’ensemble des conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-644">NOTE: The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="9b9e6-645">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-646">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-646">Parameters:</span></span>

|<span data-ttu-id="9b9e6-647">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-647">Name</span></span>| <span data-ttu-id="9b9e6-648">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-648">Type</span></span>| <span data-ttu-id="9b9e6-649">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9b9e6-650">String | Object</span><span class="sxs-lookup"><span data-stu-id="9b9e6-650">String &#124; Object</span></span>| | <span data-ttu-id="9b9e6-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9b9e6-653">**OU**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-653">**OR**</span></span><br/><span data-ttu-id="9b9e6-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9b9e6-656">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-656">String</span></span> | <span data-ttu-id="9b9e6-657">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-657">&lt;optional&gt;</span></span> | <span data-ttu-id="9b9e6-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="9b9e6-660">function</span><span class="sxs-lookup"><span data-stu-id="9b9e6-660">function</span></span> | <span data-ttu-id="9b9e6-661">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-661">&lt;optional&gt;</span></span> | <span data-ttu-id="9b9e6-662">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9b9e6-663">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-663">Requirements</span></span>

|<span data-ttu-id="9b9e6-664">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-664">Requirement</span></span>| <span data-ttu-id="9b9e6-665">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-666">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-666">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-667">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-667">1.0</span></span>|
|[<span data-ttu-id="9b9e6-668">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-669">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-670">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-671">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9b9e6-672">Exemples</span><span class="sxs-lookup"><span data-stu-id="9b9e6-672">Examples</span></span>

<span data-ttu-id="9b9e6-673">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9b9e6-674">Réponse sans texte.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9b9e6-675">Réponse avec seulement un corps de message.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9b9e6-676">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="9b9e6-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9b9e6-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="9b9e6-678">Obtient les entités figurant dans le texte de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-678">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-679">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-679">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-680">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-680">Requirements</span></span>

|<span data-ttu-id="9b9e6-681">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-681">Requirement</span></span>| <span data-ttu-id="9b9e6-682">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-683">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-684">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-684">1.0</span></span>|
|[<span data-ttu-id="9b9e6-685">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-686">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-687">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-688">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9b9e6-689">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-689">Returns:</span></span>

<span data-ttu-id="9b9e6-690">Type : [Entités](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9b9e6-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9b9e6-691">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-691">Example</span></span>

<span data-ttu-id="9b9e6-692">L’exemple suivant accède aux entités de contacts dans l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-692">The following example accesses the contacts entities on the current item.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="9b9e6-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9b9e6-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9b9e6-694">Obtient un tableau de toutes les entités du type spécifié trouvées dans le texte de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-694">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-695">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-695">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-696">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-696">Parameters:</span></span>

|<span data-ttu-id="9b9e6-697">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-697">Name</span></span>| <span data-ttu-id="9b9e6-698">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-698">Type</span></span>| <span data-ttu-id="9b9e6-699">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9b9e6-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9b9e6-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="9b9e6-701">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b9e6-702">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-702">Requirements</span></span>

|<span data-ttu-id="9b9e6-703">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-703">Requirement</span></span>| <span data-ttu-id="9b9e6-704">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-705">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-705">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-706">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-706">1.0</span></span>|
|[<span data-ttu-id="9b9e6-707">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-708">Restreint</span><span class="sxs-lookup"><span data-stu-id="9b9e6-708">Restricted</span></span>|
|[<span data-ttu-id="9b9e6-709">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-710">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9b9e6-711">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-711">Returns:</span></span>

<span data-ttu-id="9b9e6-712">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9b9e6-713">Si aucune entité du type spécifié n’est présente dans le texte de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-713">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="9b9e6-714">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9b9e6-715">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9b9e6-716">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="9b9e6-716">Value of `entityType`</span></span> | <span data-ttu-id="9b9e6-717">Type des objets dans le tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="9b9e6-717">Type of objects in returned array</span></span> | <span data-ttu-id="9b9e6-718">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="9b9e6-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9b9e6-719">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-719">String</span></span> | <span data-ttu-id="9b9e6-720">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9b9e6-721">Contact</span><span class="sxs-lookup"><span data-stu-id="9b9e6-721">Contact</span></span> | <span data-ttu-id="9b9e6-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9b9e6-723">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-723">String</span></span> | <span data-ttu-id="9b9e6-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9b9e6-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9b9e6-725">MeetingSuggestion</span></span> | <span data-ttu-id="9b9e6-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9b9e6-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9b9e6-727">PhoneNumber</span></span> | <span data-ttu-id="9b9e6-728">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9b9e6-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9b9e6-729">TaskSuggestion</span></span> | <span data-ttu-id="9b9e6-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9b9e6-731">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-731">String</span></span> | <span data-ttu-id="9b9e6-732">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="9b9e6-732">**Restricted**</span></span> |

<span data-ttu-id="9b9e6-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9b9e6-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="9b9e6-734">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-734">Example</span></span>

<span data-ttu-id="9b9e6-735">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le texte de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-735">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="9b9e6-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9b9e6-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9b9e6-737">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-738">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-738">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9b9e6-739">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-740">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-740">Parameters:</span></span>

|<span data-ttu-id="9b9e6-741">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-741">Name</span></span>| <span data-ttu-id="9b9e6-742">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-742">Type</span></span>| <span data-ttu-id="9b9e6-743">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9b9e6-744">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-744">String</span></span>|<span data-ttu-id="9b9e6-745">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b9e6-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-746">Requirements</span></span>

|<span data-ttu-id="9b9e6-747">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-747">Requirement</span></span>| <span data-ttu-id="9b9e6-748">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-749">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-750">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-750">1.0</span></span>|
|[<span data-ttu-id="9b9e6-751">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-752">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-753">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-754">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9b9e6-755">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-755">Returns:</span></span>

<span data-ttu-id="9b9e6-p146">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="9b9e6-758">Type : Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9b9e6-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="9b9e6-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9b9e6-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9b9e6-760">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-761">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-761">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9b9e6-p147">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier XML de manifeste. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9b9e6-765">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9b9e6-766">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="9b9e6-p148">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b9e6-769">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-769">Requirements</span></span>

|<span data-ttu-id="9b9e6-770">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-770">Requirement</span></span>| <span data-ttu-id="9b9e6-771">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-772">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-773">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-773">1.0</span></span>|
|[<span data-ttu-id="9b9e6-774">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-775">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-776">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-777">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9b9e6-778">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-778">Returns:</span></span>

<span data-ttu-id="9b9e6-p149">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier XML de manifeste. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9b9e6-781">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="9b9e6-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9b9e6-782">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9b9e6-783">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-783">Example</span></span>

<span data-ttu-id="9b9e6-784">L’exemple suivant montre comment accéder au tableau de correspondances pour les <rule>éléments d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="9b9e6-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9b9e6-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="9b9e6-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9b9e6-786">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier XML de manifeste.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9b9e6-787">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-787">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9b9e6-788">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier XML de manifeste ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9b9e6-p150">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de texte d’un élément, l’expression régulière doit filtrer davantage le texte. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du texte de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du texte d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-791">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-791">Parameters:</span></span>

|<span data-ttu-id="9b9e6-792">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-792">Name</span></span>| <span data-ttu-id="9b9e6-793">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-793">Type</span></span>| <span data-ttu-id="9b9e6-794">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9b9e6-795">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-795">String</span></span>|<span data-ttu-id="9b9e6-796">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b9e6-797">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-797">Requirements</span></span>

|<span data-ttu-id="9b9e6-798">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-798">Requirement</span></span>| <span data-ttu-id="9b9e6-799">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-800">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-800">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-801">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-801">1.0</span></span>|
|[<span data-ttu-id="9b9e6-802">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-803">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-804">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-805">Lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9b9e6-806">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-806">Returns:</span></span>

<span data-ttu-id="9b9e6-807">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9b9e6-808">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="9b9e6-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9b9e6-809">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9b9e6-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9b9e6-810">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9b9e6-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9b9e6-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9b9e6-812">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9b9e6-p151">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-816">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-816">Parameters:</span></span>

|<span data-ttu-id="9b9e6-817">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-817">Name</span></span>| <span data-ttu-id="9b9e6-818">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-818">Type</span></span>| <span data-ttu-id="9b9e6-819">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-819">Attributes</span></span>| <span data-ttu-id="9b9e6-820">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9b9e6-821">function</span><span class="sxs-lookup"><span data-stu-id="9b9e6-821">function</span></span>||<span data-ttu-id="9b9e6-822">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9b9e6-823">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9b9e6-824">Cet objet peut être utilisé pour obtenir, définir et supprimer les propriétés personnalisées de l’élément et sauvegarder les modifications du jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-824">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9b9e6-825">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-825">Object</span></span>| <span data-ttu-id="9b9e6-826">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-826">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-827">Les développeurs peuvent fournir n'importe quel objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-827">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="9b9e6-828">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b9e6-829">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-829">Requirements</span></span>

|<span data-ttu-id="9b9e6-830">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-830">Requirement</span></span>| <span data-ttu-id="9b9e6-831">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-832">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-832">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-833">1.0</span><span class="sxs-lookup"><span data-stu-id="9b9e6-833">1.0</span></span>|
|[<span data-ttu-id="9b9e6-834">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-835">ReadItem</span></span>|
|[<span data-ttu-id="9b9e6-836">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-837">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b9e6-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-838">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-838">Example</span></span>

<span data-ttu-id="9b9e6-p154">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9b9e6-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9b9e6-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9b9e6-843">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9b9e6-p155">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les appareils, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire inclus et qu’il le fait ensuite apparaître dans une nouvelle fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9b9e6-848">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9b9e6-848">Parameters:</span></span>

|<span data-ttu-id="9b9e6-849">Nom</span><span class="sxs-lookup"><span data-stu-id="9b9e6-849">Name</span></span>| <span data-ttu-id="9b9e6-850">Type</span><span class="sxs-lookup"><span data-stu-id="9b9e6-850">Type</span></span>| <span data-ttu-id="9b9e6-851">Attributs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-851">Attributes</span></span>| <span data-ttu-id="9b9e6-852">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9b9e6-853">String</span><span class="sxs-lookup"><span data-stu-id="9b9e6-853">String</span></span>||<span data-ttu-id="9b9e6-p156">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="9b9e6-856">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-856">Object</span></span>| <span data-ttu-id="9b9e6-857">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-857">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-858">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9b9e6-859">Objet</span><span class="sxs-lookup"><span data-stu-id="9b9e6-859">Object</span></span>| <span data-ttu-id="9b9e6-860">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-860">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-861">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9b9e6-862">function</span><span class="sxs-lookup"><span data-stu-id="9b9e6-862">function</span></span>| <span data-ttu-id="9b9e6-863">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="9b9e6-863">&lt;optional&gt;</span></span>|<span data-ttu-id="9b9e6-864">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9b9e6-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9b9e6-865">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9b9e6-866">Erreurs</span><span class="sxs-lookup"><span data-stu-id="9b9e6-866">Errors</span></span>

| <span data-ttu-id="9b9e6-867">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-867">Error code</span></span> | <span data-ttu-id="9b9e6-868">Description</span><span class="sxs-lookup"><span data-stu-id="9b9e6-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9b9e6-869">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="9b9e6-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9b9e6-870">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b9e6-870">Requirements</span></span>

|<span data-ttu-id="9b9e6-871">Condition requise</span><span class="sxs-lookup"><span data-stu-id="9b9e6-871">Requirement</span></span>| <span data-ttu-id="9b9e6-872">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b9e6-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b9e6-873">Version minimale de l’ensemble de conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b9e6-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9b9e6-874">1.1</span><span class="sxs-lookup"><span data-stu-id="9b9e6-874">1.1</span></span>|
|[<span data-ttu-id="9b9e6-875">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="9b9e6-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9b9e6-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9b9e6-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="9b9e6-877">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b9e6-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9b9e6-878">Composition</span><span class="sxs-lookup"><span data-stu-id="9b9e6-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9b9e6-879">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b9e6-879">Example</span></span>

<span data-ttu-id="9b9e6-880">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="9b9e6-880">The following code removes an attachment with an identifier of '0'.</span></span>

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