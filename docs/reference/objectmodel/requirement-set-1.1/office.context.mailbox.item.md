---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d3681f369570995c07256171fb6a65482648e85e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450379"
---
# <a name="item"></a><span data-ttu-id="97607-102">élément</span><span class="sxs-lookup"><span data-stu-id="97607-102">item</span></span>

### <span data-ttu-id="97607-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="97607-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="97607-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="97607-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-107">Requirements</span></span>

|<span data-ttu-id="97607-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-108">Requirement</span></span>| <span data-ttu-id="97607-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-111">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-111">1.0</span></span>|
|[<span data-ttu-id="97607-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="97607-113">Restricted</span></span>|
|[<span data-ttu-id="97607-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-115">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="97607-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-116">Example</span></span>

<span data-ttu-id="97607-117">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="97607-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="97607-118">Membres</span><span class="sxs-lookup"><span data-stu-id="97607-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="97607-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="97607-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="97607-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-122">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="97607-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="97607-123">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="97607-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="97607-124">Type</span><span class="sxs-lookup"><span data-stu-id="97607-124">Type</span></span>

*   <span data-ttu-id="97607-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="97607-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-126">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-126">Requirements</span></span>

|<span data-ttu-id="97607-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-127">Requirement</span></span>| <span data-ttu-id="97607-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-129">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-130">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-130">1.0</span></span>|
|[<span data-ttu-id="97607-131">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-132">ReadItem</span></span>|
|[<span data-ttu-id="97607-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-134">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-135">Example</span></span>

<span data-ttu-id="97607-136">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="97607-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="97607-138">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="97607-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="97607-139">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-140">Type</span><span class="sxs-lookup"><span data-stu-id="97607-140">Type</span></span>

*   [<span data-ttu-id="97607-141">Destinataires</span><span class="sxs-lookup"><span data-stu-id="97607-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="97607-142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-142">Requirements</span></span>

|<span data-ttu-id="97607-143">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-143">Requirement</span></span>| <span data-ttu-id="97607-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-145">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-146">1.1</span><span class="sxs-lookup"><span data-stu-id="97607-146">1.1</span></span>|
|[<span data-ttu-id="97607-147">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-148">ReadItem</span></span>|
|[<span data-ttu-id="97607-149">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-150">Composition</span><span class="sxs-lookup"><span data-stu-id="97607-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-151">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="97607-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="97607-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="97607-153">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="97607-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-154">Type</span><span class="sxs-lookup"><span data-stu-id="97607-154">Type</span></span>

*   [<span data-ttu-id="97607-155">Body</span><span class="sxs-lookup"><span data-stu-id="97607-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="97607-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-156">Requirements</span></span>

|<span data-ttu-id="97607-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-157">Requirement</span></span>| <span data-ttu-id="97607-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-160">1.1</span><span class="sxs-lookup"><span data-stu-id="97607-160">1.1</span></span>|
|[<span data-ttu-id="97607-161">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-162">ReadItem</span></span>|
|[<span data-ttu-id="97607-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-165">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-165">Example</span></span>

<span data-ttu-id="97607-166">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="97607-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="97607-167">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="97607-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="97607-169">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="97607-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="97607-170">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-171">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-171">Read mode</span></span>

<span data-ttu-id="97607-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="97607-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-174">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-174">Compose mode</span></span>

<span data-ttu-id="97607-175">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="97607-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97607-176">Type</span><span class="sxs-lookup"><span data-stu-id="97607-176">Type</span></span>

*   <span data-ttu-id="97607-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-178">Requirements</span></span>

|<span data-ttu-id="97607-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-179">Requirement</span></span>| <span data-ttu-id="97607-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-182">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-182">1.0</span></span>|
|[<span data-ttu-id="97607-183">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-184">ReadItem</span></span>|
|[<span data-ttu-id="97607-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="97607-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="97607-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="97607-188">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="97607-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="97607-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="97607-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="97607-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="97607-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-193">Type</span><span class="sxs-lookup"><span data-stu-id="97607-193">Type</span></span>

*   <span data-ttu-id="97607-194">String</span><span class="sxs-lookup"><span data-stu-id="97607-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-195">Requirements</span></span>

|<span data-ttu-id="97607-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-196">Requirement</span></span>| <span data-ttu-id="97607-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-199">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-199">1.0</span></span>|
|[<span data-ttu-id="97607-200">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-201">ReadItem</span></span>|
|[<span data-ttu-id="97607-202">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-203">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-204">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="97607-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="97607-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="97607-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-208">Type</span><span class="sxs-lookup"><span data-stu-id="97607-208">Type</span></span>

*   <span data-ttu-id="97607-209">Date</span><span class="sxs-lookup"><span data-stu-id="97607-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-210">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-210">Requirements</span></span>

|<span data-ttu-id="97607-211">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-211">Requirement</span></span>| <span data-ttu-id="97607-212">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-213">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-214">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-214">1.0</span></span>|
|[<span data-ttu-id="97607-215">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-216">ReadItem</span></span>|
|[<span data-ttu-id="97607-217">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-218">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-219">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="97607-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="97607-220">dateTimeModified :Date</span></span>

<span data-ttu-id="97607-p111">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-223">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-224">Type</span><span class="sxs-lookup"><span data-stu-id="97607-224">Type</span></span>

*   <span data-ttu-id="97607-225">Date</span><span class="sxs-lookup"><span data-stu-id="97607-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-226">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-226">Requirements</span></span>

|<span data-ttu-id="97607-227">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-227">Requirement</span></span>| <span data-ttu-id="97607-228">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-229">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-230">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-230">1.0</span></span>|
|[<span data-ttu-id="97607-231">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-232">ReadItem</span></span>|
|[<span data-ttu-id="97607-233">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-234">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-235">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="97607-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="97607-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="97607-237">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="97607-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="97607-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-240">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="97607-240">Read mode</span></span>

<span data-ttu-id="97607-241">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="97607-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-242">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-242">Compose mode</span></span>

<span data-ttu-id="97607-243">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="97607-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="97607-244">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="97607-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="97607-245">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="97607-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="97607-246">Type</span><span class="sxs-lookup"><span data-stu-id="97607-246">Type</span></span>

*   <span data-ttu-id="97607-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="97607-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-248">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-248">Requirements</span></span>

|<span data-ttu-id="97607-249">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-249">Requirement</span></span>| <span data-ttu-id="97607-250">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-251">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-252">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-252">1.0</span></span>|
|[<span data-ttu-id="97607-253">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-254">ReadItem</span></span>|
|[<span data-ttu-id="97607-255">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-256">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="97607-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="97607-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="97607-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="97607-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="97607-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-262">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="97607-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-263">Type</span><span class="sxs-lookup"><span data-stu-id="97607-263">Type</span></span>

*   [<span data-ttu-id="97607-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="97607-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="97607-265">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-265">Requirements</span></span>

|<span data-ttu-id="97607-266">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-266">Requirement</span></span>| <span data-ttu-id="97607-267">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-268">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-269">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-269">1.0</span></span>|
|[<span data-ttu-id="97607-270">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-271">ReadItem</span></span>|
|[<span data-ttu-id="97607-272">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-273">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-274">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="97607-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="97607-275">internetMessageId :String</span></span>

<span data-ttu-id="97607-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-278">Type</span><span class="sxs-lookup"><span data-stu-id="97607-278">Type</span></span>

*   <span data-ttu-id="97607-279">String</span><span class="sxs-lookup"><span data-stu-id="97607-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-280">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-280">Requirements</span></span>

|<span data-ttu-id="97607-281">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-281">Requirement</span></span>| <span data-ttu-id="97607-282">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-283">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-284">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-284">1.0</span></span>|
|[<span data-ttu-id="97607-285">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-286">ReadItem</span></span>|
|[<span data-ttu-id="97607-287">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-288">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-289">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="97607-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="97607-290">itemClass :String</span></span>

<span data-ttu-id="97607-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="97607-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="97607-295">Type</span><span class="sxs-lookup"><span data-stu-id="97607-295">Type</span></span> | <span data-ttu-id="97607-296">Description</span><span class="sxs-lookup"><span data-stu-id="97607-296">Description</span></span> | <span data-ttu-id="97607-297">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="97607-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="97607-298">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="97607-298">Appointment items</span></span> | <span data-ttu-id="97607-299">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="97607-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="97607-300">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="97607-300">Message items</span></span> | <span data-ttu-id="97607-301">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="97607-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="97607-302">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="97607-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-303">Type</span><span class="sxs-lookup"><span data-stu-id="97607-303">Type</span></span>

*   <span data-ttu-id="97607-304">String</span><span class="sxs-lookup"><span data-stu-id="97607-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-305">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-305">Requirements</span></span>

|<span data-ttu-id="97607-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-306">Requirement</span></span>| <span data-ttu-id="97607-307">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-308">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-309">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-309">1.0</span></span>|
|[<span data-ttu-id="97607-310">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-311">ReadItem</span></span>|
|[<span data-ttu-id="97607-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-313">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-314">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="97607-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="97607-315">(nullable) itemId :String</span></span>

<span data-ttu-id="97607-p118">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-318">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="97607-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="97607-319">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="97607-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="97607-320">Avant d'effectuer des appels d'API REST à l'aide de cette valeur `Office.context.mailbox.convertToRestId`, elle doit être convertie à l'aide de, qui est disponible à partir de l'ensemble de conditions requises 1,3.</span><span class="sxs-lookup"><span data-stu-id="97607-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="97607-321">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="97607-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="97607-322">Type</span><span class="sxs-lookup"><span data-stu-id="97607-322">Type</span></span>

*   <span data-ttu-id="97607-323">String</span><span class="sxs-lookup"><span data-stu-id="97607-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-324">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-324">Requirements</span></span>

|<span data-ttu-id="97607-325">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-325">Requirement</span></span>| <span data-ttu-id="97607-326">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-327">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-328">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-328">1.0</span></span>|
|[<span data-ttu-id="97607-329">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-330">ReadItem</span></span>|
|[<span data-ttu-id="97607-331">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-332">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-333">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-333">Example</span></span>

<span data-ttu-id="97607-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="97607-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="97607-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="97607-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="97607-337">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="97607-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="97607-338">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-339">Type</span><span class="sxs-lookup"><span data-stu-id="97607-339">Type</span></span>

*   [<span data-ttu-id="97607-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="97607-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="97607-341">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-341">Requirements</span></span>

|<span data-ttu-id="97607-342">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-342">Requirement</span></span>| <span data-ttu-id="97607-343">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-344">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-345">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-345">1.0</span></span>|
|[<span data-ttu-id="97607-346">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-347">ReadItem</span></span>|
|[<span data-ttu-id="97607-348">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-349">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-350">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="97607-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="97607-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="97607-352">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-353">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="97607-353">Read mode</span></span>

<span data-ttu-id="97607-354">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-355">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-355">Compose mode</span></span>

<span data-ttu-id="97607-356">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97607-357">Type</span><span class="sxs-lookup"><span data-stu-id="97607-357">Type</span></span>

*   <span data-ttu-id="97607-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="97607-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-359">Requirements</span></span>

|<span data-ttu-id="97607-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-360">Requirement</span></span>| <span data-ttu-id="97607-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-363">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-363">1.0</span></span>|
|[<span data-ttu-id="97607-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-365">ReadItem</span></span>|
|[<span data-ttu-id="97607-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-367">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="97607-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="97607-368">normalizedSubject :String</span></span>

<span data-ttu-id="97607-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="97607-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="97607-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-373">Type</span><span class="sxs-lookup"><span data-stu-id="97607-373">Type</span></span>

*   <span data-ttu-id="97607-374">String</span><span class="sxs-lookup"><span data-stu-id="97607-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-375">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-375">Requirements</span></span>

|<span data-ttu-id="97607-376">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-376">Requirement</span></span>| <span data-ttu-id="97607-377">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-378">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-379">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-379">1.0</span></span>|
|[<span data-ttu-id="97607-380">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-381">ReadItem</span></span>|
|[<span data-ttu-id="97607-382">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-383">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-384">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="97607-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="97607-386">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="97607-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="97607-387">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-388">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-388">Read mode</span></span>

<span data-ttu-id="97607-389">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="97607-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-390">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-390">Compose mode</span></span>

<span data-ttu-id="97607-391">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="97607-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97607-392">Type</span><span class="sxs-lookup"><span data-stu-id="97607-392">Type</span></span>

*   <span data-ttu-id="97607-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-394">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-394">Requirements</span></span>

|<span data-ttu-id="97607-395">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-395">Requirement</span></span>| <span data-ttu-id="97607-396">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-397">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-398">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-398">1.0</span></span>|
|[<span data-ttu-id="97607-399">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-400">ReadItem</span></span>|
|[<span data-ttu-id="97607-401">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-402">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="97607-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="97607-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="97607-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-406">Type</span><span class="sxs-lookup"><span data-stu-id="97607-406">Type</span></span>

*   [<span data-ttu-id="97607-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="97607-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="97607-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-408">Requirements</span></span>

|<span data-ttu-id="97607-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-409">Requirement</span></span>| <span data-ttu-id="97607-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-412">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-412">1.0</span></span>|
|[<span data-ttu-id="97607-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-414">ReadItem</span></span>|
|[<span data-ttu-id="97607-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="97607-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="97607-419">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="97607-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="97607-420">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-421">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-421">Read mode</span></span>

<span data-ttu-id="97607-422">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="97607-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-423">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-423">Compose mode</span></span>

<span data-ttu-id="97607-424">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="97607-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="97607-425">Type</span><span class="sxs-lookup"><span data-stu-id="97607-425">Type</span></span>

*   <span data-ttu-id="97607-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-427">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-427">Requirements</span></span>

|<span data-ttu-id="97607-428">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-428">Requirement</span></span>| <span data-ttu-id="97607-429">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-430">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-431">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-431">1.0</span></span>|
|[<span data-ttu-id="97607-432">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-433">ReadItem</span></span>|
|[<span data-ttu-id="97607-434">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-435">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="97607-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="97607-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="97607-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="97607-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="97607-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="97607-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-441">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="97607-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="97607-442">Type</span><span class="sxs-lookup"><span data-stu-id="97607-442">Type</span></span>

*   [<span data-ttu-id="97607-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="97607-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="97607-444">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-444">Requirements</span></span>

|<span data-ttu-id="97607-445">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-445">Requirement</span></span>| <span data-ttu-id="97607-446">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-447">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-448">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-448">1.0</span></span>|
|[<span data-ttu-id="97607-449">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-450">ReadItem</span></span>|
|[<span data-ttu-id="97607-451">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-452">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-453">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="97607-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="97607-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="97607-455">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="97607-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="97607-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-458">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="97607-458">Read mode</span></span>

<span data-ttu-id="97607-459">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="97607-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-460">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-460">Compose mode</span></span>

<span data-ttu-id="97607-461">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="97607-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="97607-462">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="97607-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="97607-463">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="97607-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="97607-464">Type</span><span class="sxs-lookup"><span data-stu-id="97607-464">Type</span></span>

*   <span data-ttu-id="97607-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="97607-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-466">Requirements</span></span>

|<span data-ttu-id="97607-467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-467">Requirement</span></span>| <span data-ttu-id="97607-468">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-470">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-470">1.0</span></span>|
|[<span data-ttu-id="97607-471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-472">ReadItem</span></span>|
|[<span data-ttu-id="97607-473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-474">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="97607-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="97607-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="97607-476">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="97607-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="97607-477">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="97607-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-478">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="97607-478">Read mode</span></span>

<span data-ttu-id="97607-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="97607-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-481">Compose mode</span></span>

<span data-ttu-id="97607-482">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="97607-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="97607-483">Type</span><span class="sxs-lookup"><span data-stu-id="97607-483">Type</span></span>

*   <span data-ttu-id="97607-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="97607-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-485">Requirements</span></span>

|<span data-ttu-id="97607-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-486">Requirement</span></span>| <span data-ttu-id="97607-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-489">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-489">1.0</span></span>|
|[<span data-ttu-id="97607-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-491">ReadItem</span></span>|
|[<span data-ttu-id="97607-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-493">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="97607-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="97607-495">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="97607-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="97607-496">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97607-497">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-497">Read mode</span></span>

<span data-ttu-id="97607-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="97607-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="97607-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="97607-500">Compose mode</span></span>

<span data-ttu-id="97607-501">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="97607-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97607-502">Type</span><span class="sxs-lookup"><span data-stu-id="97607-502">Type</span></span>

*   <span data-ttu-id="97607-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="97607-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-504">Requirements</span></span>

|<span data-ttu-id="97607-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-505">Requirement</span></span>| <span data-ttu-id="97607-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-508">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-508">1.0</span></span>|
|[<span data-ttu-id="97607-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-510">ReadItem</span></span>|
|[<span data-ttu-id="97607-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-512">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="97607-513">Méthodes</span><span class="sxs-lookup"><span data-stu-id="97607-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="97607-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="97607-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="97607-515">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="97607-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="97607-516">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="97607-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="97607-517">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="97607-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-518">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-518">Parameters</span></span>

|<span data-ttu-id="97607-519">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-519">Name</span></span>| <span data-ttu-id="97607-520">Type</span><span class="sxs-lookup"><span data-stu-id="97607-520">Type</span></span>| <span data-ttu-id="97607-521">Attributs</span><span class="sxs-lookup"><span data-stu-id="97607-521">Attributes</span></span>| <span data-ttu-id="97607-522">Description</span><span class="sxs-lookup"><span data-stu-id="97607-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="97607-523">String</span><span class="sxs-lookup"><span data-stu-id="97607-523">String</span></span>||<span data-ttu-id="97607-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="97607-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="97607-526">String</span><span class="sxs-lookup"><span data-stu-id="97607-526">String</span></span>||<span data-ttu-id="97607-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="97607-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="97607-529">Objet</span><span class="sxs-lookup"><span data-stu-id="97607-529">Object</span></span>| <span data-ttu-id="97607-530">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-530">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-531">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="97607-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97607-532">Objet</span><span class="sxs-lookup"><span data-stu-id="97607-532">Object</span></span>| <span data-ttu-id="97607-533">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-533">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-534">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97607-535">fonction</span><span class="sxs-lookup"><span data-stu-id="97607-535">function</span></span>| <span data-ttu-id="97607-536">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-536">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-537">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="97607-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="97607-538">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="97607-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="97607-539">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="97607-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="97607-540">Erreurs</span><span class="sxs-lookup"><span data-stu-id="97607-540">Errors</span></span>

| <span data-ttu-id="97607-541">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="97607-541">Error code</span></span> | <span data-ttu-id="97607-542">Description</span><span class="sxs-lookup"><span data-stu-id="97607-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="97607-543">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="97607-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="97607-544">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="97607-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="97607-545">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="97607-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97607-546">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-546">Requirements</span></span>

|<span data-ttu-id="97607-547">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-547">Requirement</span></span>| <span data-ttu-id="97607-548">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-549">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-550">1.1</span><span class="sxs-lookup"><span data-stu-id="97607-550">1.1</span></span>|
|[<span data-ttu-id="97607-551">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97607-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="97607-553">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-554">Composition</span><span class="sxs-lookup"><span data-stu-id="97607-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-555">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="97607-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="97607-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="97607-557">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="97607-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="97607-561">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="97607-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="97607-562">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="97607-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-563">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-563">Parameters</span></span>

|<span data-ttu-id="97607-564">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-564">Name</span></span>| <span data-ttu-id="97607-565">Type</span><span class="sxs-lookup"><span data-stu-id="97607-565">Type</span></span>| <span data-ttu-id="97607-566">Attributs</span><span class="sxs-lookup"><span data-stu-id="97607-566">Attributes</span></span>| <span data-ttu-id="97607-567">Description</span><span class="sxs-lookup"><span data-stu-id="97607-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="97607-568">String</span><span class="sxs-lookup"><span data-stu-id="97607-568">String</span></span>||<span data-ttu-id="97607-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="97607-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="97607-571">String</span><span class="sxs-lookup"><span data-stu-id="97607-571">String</span></span>||<span data-ttu-id="97607-572">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="97607-572">The subject of the item to be attached.</span></span> <span data-ttu-id="97607-573">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="97607-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="97607-574">Object</span><span class="sxs-lookup"><span data-stu-id="97607-574">Object</span></span>| <span data-ttu-id="97607-575">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-575">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-576">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="97607-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97607-577">Objet</span><span class="sxs-lookup"><span data-stu-id="97607-577">Object</span></span>| <span data-ttu-id="97607-578">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-578">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-579">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97607-580">fonction</span><span class="sxs-lookup"><span data-stu-id="97607-580">function</span></span>| <span data-ttu-id="97607-581">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-581">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-582">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="97607-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="97607-583">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="97607-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="97607-584">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="97607-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="97607-585">Erreurs</span><span class="sxs-lookup"><span data-stu-id="97607-585">Errors</span></span>

| <span data-ttu-id="97607-586">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="97607-586">Error code</span></span> | <span data-ttu-id="97607-587">Description</span><span class="sxs-lookup"><span data-stu-id="97607-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="97607-588">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="97607-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97607-589">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-589">Requirements</span></span>

|<span data-ttu-id="97607-590">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-590">Requirement</span></span>| <span data-ttu-id="97607-591">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-592">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-593">1.1</span><span class="sxs-lookup"><span data-stu-id="97607-593">1.1</span></span>|
|[<span data-ttu-id="97607-594">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97607-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="97607-596">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-597">Composition</span><span class="sxs-lookup"><span data-stu-id="97607-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-598">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-598">Example</span></span>

<span data-ttu-id="97607-599">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="97607-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="97607-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="97607-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="97607-601">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="97607-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-602">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="97607-603">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="97607-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="97607-604">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="97607-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-605">La possibilité d'inclure des pièces jointes dans `displayReplyAllForm` l'appel à n'est pas prise en charge dans l'ensemble de conditions requises 1,1.</span><span class="sxs-lookup"><span data-stu-id="97607-605">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="97607-606">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="97607-606">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-607">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-607">Parameters</span></span>

|<span data-ttu-id="97607-608">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-608">Name</span></span>| <span data-ttu-id="97607-609">Type</span><span class="sxs-lookup"><span data-stu-id="97607-609">Type</span></span>| <span data-ttu-id="97607-610">Description</span><span class="sxs-lookup"><span data-stu-id="97607-610">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="97607-611">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="97607-611">String &#124; Object</span></span>| |<span data-ttu-id="97607-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="97607-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="97607-614">**OU**</span><span class="sxs-lookup"><span data-stu-id="97607-614">**OR**</span></span><br/><span data-ttu-id="97607-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="97607-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="97607-617">String</span><span class="sxs-lookup"><span data-stu-id="97607-617">String</span></span> | <span data-ttu-id="97607-618">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-618">&lt;optional&gt;</span></span> | <span data-ttu-id="97607-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="97607-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="97607-621">fonction</span><span class="sxs-lookup"><span data-stu-id="97607-621">function</span></span> | <span data-ttu-id="97607-622">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-622">&lt;optional&gt;</span></span> | <span data-ttu-id="97607-623">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="97607-623">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97607-624">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-624">Requirements</span></span>

|<span data-ttu-id="97607-625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-625">Requirement</span></span>| <span data-ttu-id="97607-626">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-627">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-628">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-628">1.0</span></span>|
|[<span data-ttu-id="97607-629">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-630">ReadItem</span></span>|
|[<span data-ttu-id="97607-631">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-632">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-632">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="97607-633">Exemples</span><span class="sxs-lookup"><span data-stu-id="97607-633">Examples</span></span>

<span data-ttu-id="97607-634">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="97607-634">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="97607-635">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="97607-635">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="97607-636">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="97607-636">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="97607-637">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-637">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="97607-638">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="97607-638">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="97607-639">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="97607-639">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-640">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-640">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="97607-641">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="97607-641">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="97607-642">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="97607-642">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-643">La possibilité d'inclure des pièces jointes dans `displayReplyForm` l'appel à n'est pas prise en charge dans l'ensemble de conditions requises 1,1.</span><span class="sxs-lookup"><span data-stu-id="97607-643">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="97607-644">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="97607-644">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-645">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-645">Parameters</span></span>

|<span data-ttu-id="97607-646">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-646">Name</span></span>| <span data-ttu-id="97607-647">Type</span><span class="sxs-lookup"><span data-stu-id="97607-647">Type</span></span>| <span data-ttu-id="97607-648">Description</span><span class="sxs-lookup"><span data-stu-id="97607-648">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="97607-649">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="97607-649">String &#124; Object</span></span>| | <span data-ttu-id="97607-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="97607-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="97607-652">**OU**</span><span class="sxs-lookup"><span data-stu-id="97607-652">**OR**</span></span><br/><span data-ttu-id="97607-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="97607-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="97607-655">String</span><span class="sxs-lookup"><span data-stu-id="97607-655">String</span></span> | <span data-ttu-id="97607-656">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-656">&lt;optional&gt;</span></span> | <span data-ttu-id="97607-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="97607-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="97607-659">fonction</span><span class="sxs-lookup"><span data-stu-id="97607-659">function</span></span> | <span data-ttu-id="97607-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-660">&lt;optional&gt;</span></span> | <span data-ttu-id="97607-661">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="97607-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97607-662">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-662">Requirements</span></span>

|<span data-ttu-id="97607-663">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-663">Requirement</span></span>| <span data-ttu-id="97607-664">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-665">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-666">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-666">1.0</span></span>|
|[<span data-ttu-id="97607-667">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-668">ReadItem</span></span>|
|[<span data-ttu-id="97607-669">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-670">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-670">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="97607-671">Exemples</span><span class="sxs-lookup"><span data-stu-id="97607-671">Examples</span></span>

<span data-ttu-id="97607-672">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="97607-672">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="97607-673">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="97607-673">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="97607-674">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="97607-674">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="97607-675">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-675">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="97607-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="97607-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="97607-677">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="97607-677">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-678">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-678">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-679">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-679">Requirements</span></span>

|<span data-ttu-id="97607-680">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-680">Requirement</span></span>| <span data-ttu-id="97607-681">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-681">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-682">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-682">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-683">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-683">1.0</span></span>|
|[<span data-ttu-id="97607-684">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-684">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-685">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-685">ReadItem</span></span>|
|[<span data-ttu-id="97607-686">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-686">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-687">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-687">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97607-688">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="97607-688">Returns:</span></span>

<span data-ttu-id="97607-689">Type : [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="97607-689">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="97607-690">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-690">Example</span></span>

<span data-ttu-id="97607-691">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-691">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="97607-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="97607-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="97607-693">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="97607-693">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-694">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-694">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-695">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-695">Parameters</span></span>

|<span data-ttu-id="97607-696">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-696">Name</span></span>| <span data-ttu-id="97607-697">Type</span><span class="sxs-lookup"><span data-stu-id="97607-697">Type</span></span>| <span data-ttu-id="97607-698">Description</span><span class="sxs-lookup"><span data-stu-id="97607-698">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="97607-699">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="97607-699">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="97607-700">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="97607-700">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97607-701">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-701">Requirements</span></span>

|<span data-ttu-id="97607-702">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-702">Requirement</span></span>| <span data-ttu-id="97607-703">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-703">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-704">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-704">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-705">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-705">1.0</span></span>|
|[<span data-ttu-id="97607-706">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-706">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-707">Restreinte</span><span class="sxs-lookup"><span data-stu-id="97607-707">Restricted</span></span>|
|[<span data-ttu-id="97607-708">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-708">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-709">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-709">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97607-710">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="97607-710">Returns:</span></span>

<span data-ttu-id="97607-711">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="97607-711">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="97607-712">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="97607-712">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="97607-713">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="97607-713">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="97607-714">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="97607-714">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="97607-715">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="97607-715">Value of `entityType`</span></span> | <span data-ttu-id="97607-716">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="97607-716">Type of objects in returned array</span></span> | <span data-ttu-id="97607-717">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="97607-717">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="97607-718">String</span><span class="sxs-lookup"><span data-stu-id="97607-718">String</span></span> | <span data-ttu-id="97607-719">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="97607-719">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="97607-720">Contact</span><span class="sxs-lookup"><span data-stu-id="97607-720">Contact</span></span> | <span data-ttu-id="97607-721">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97607-721">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="97607-722">String</span><span class="sxs-lookup"><span data-stu-id="97607-722">String</span></span> | <span data-ttu-id="97607-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97607-723">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="97607-724">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="97607-724">MeetingSuggestion</span></span> | <span data-ttu-id="97607-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97607-725">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="97607-726">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="97607-726">PhoneNumber</span></span> | <span data-ttu-id="97607-727">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="97607-727">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="97607-728">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="97607-728">TaskSuggestion</span></span> | <span data-ttu-id="97607-729">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97607-729">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="97607-730">String</span><span class="sxs-lookup"><span data-stu-id="97607-730">String</span></span> | <span data-ttu-id="97607-731">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="97607-731">**Restricted**</span></span> |

<span data-ttu-id="97607-732">Type :  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="97607-732">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="97607-733">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-733">Example</span></span>

<span data-ttu-id="97607-734">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="97607-734">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="97607-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="97607-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="97607-736">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="97607-736">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-737">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-737">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="97607-738">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="97607-738">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-739">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-739">Parameters</span></span>

|<span data-ttu-id="97607-740">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-740">Name</span></span>| <span data-ttu-id="97607-741">Type</span><span class="sxs-lookup"><span data-stu-id="97607-741">Type</span></span>| <span data-ttu-id="97607-742">Description</span><span class="sxs-lookup"><span data-stu-id="97607-742">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="97607-743">String</span><span class="sxs-lookup"><span data-stu-id="97607-743">String</span></span>|<span data-ttu-id="97607-744">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="97607-744">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97607-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-745">Requirements</span></span>

|<span data-ttu-id="97607-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-746">Requirement</span></span>| <span data-ttu-id="97607-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-749">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-749">1.0</span></span>|
|[<span data-ttu-id="97607-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-751">ReadItem</span></span>|
|[<span data-ttu-id="97607-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-753">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97607-754">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="97607-754">Returns:</span></span>

<span data-ttu-id="97607-p146">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="97607-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="97607-757">Type : Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="97607-757">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="97607-758">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="97607-758">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="97607-759">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="97607-759">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-760">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="97607-p147">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="97607-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="97607-764">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="97607-764">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="97607-765">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="97607-765">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="97607-p148">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="97607-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97607-768">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-768">Requirements</span></span>

|<span data-ttu-id="97607-769">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-769">Requirement</span></span>| <span data-ttu-id="97607-770">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-771">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-772">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-772">1.0</span></span>|
|[<span data-ttu-id="97607-773">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-773">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-774">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-774">ReadItem</span></span>|
|[<span data-ttu-id="97607-775">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-775">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-776">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97607-777">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="97607-777">Returns:</span></span>

<span data-ttu-id="97607-p149">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="97607-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="97607-780">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="97607-780">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="97607-781">Object</span><span class="sxs-lookup"><span data-stu-id="97607-781">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="97607-782">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-782">Example</span></span>

<span data-ttu-id="97607-783">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="97607-783">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="97607-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="97607-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="97607-785">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="97607-785">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="97607-786">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="97607-786">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="97607-787">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="97607-787">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="97607-p150">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="97607-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-790">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-790">Parameters</span></span>

|<span data-ttu-id="97607-791">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-791">Name</span></span>| <span data-ttu-id="97607-792">Type</span><span class="sxs-lookup"><span data-stu-id="97607-792">Type</span></span>| <span data-ttu-id="97607-793">Description</span><span class="sxs-lookup"><span data-stu-id="97607-793">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="97607-794">String</span><span class="sxs-lookup"><span data-stu-id="97607-794">String</span></span>|<span data-ttu-id="97607-795">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="97607-795">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97607-796">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-796">Requirements</span></span>

|<span data-ttu-id="97607-797">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-797">Requirement</span></span>| <span data-ttu-id="97607-798">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-799">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-800">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-800">1.0</span></span>|
|[<span data-ttu-id="97607-801">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-802">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-802">ReadItem</span></span>|
|[<span data-ttu-id="97607-803">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-804">Lecture</span><span class="sxs-lookup"><span data-stu-id="97607-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97607-805">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="97607-805">Returns:</span></span>

<span data-ttu-id="97607-806">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="97607-806">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="97607-807">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="97607-807">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="97607-808">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="97607-808">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="97607-809">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-809">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="97607-810">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="97607-810">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="97607-811">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="97607-811">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="97607-p151">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="97607-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-815">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-815">Parameters</span></span>

|<span data-ttu-id="97607-816">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-816">Name</span></span>| <span data-ttu-id="97607-817">Type</span><span class="sxs-lookup"><span data-stu-id="97607-817">Type</span></span>| <span data-ttu-id="97607-818">Attributs</span><span class="sxs-lookup"><span data-stu-id="97607-818">Attributes</span></span>| <span data-ttu-id="97607-819">Description</span><span class="sxs-lookup"><span data-stu-id="97607-819">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="97607-820">function</span><span class="sxs-lookup"><span data-stu-id="97607-820">function</span></span>||<span data-ttu-id="97607-821">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="97607-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="97607-822">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="97607-822">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="97607-823">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="97607-823">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="97607-824">Objet</span><span class="sxs-lookup"><span data-stu-id="97607-824">Object</span></span>| <span data-ttu-id="97607-825">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-825">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-826">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-826">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="97607-827">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-827">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97607-828">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-828">Requirements</span></span>

|<span data-ttu-id="97607-829">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-829">Requirement</span></span>| <span data-ttu-id="97607-830">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-830">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-831">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-831">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-832">1.0</span><span class="sxs-lookup"><span data-stu-id="97607-832">1.0</span></span>|
|[<span data-ttu-id="97607-833">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-833">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-834">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97607-834">ReadItem</span></span>|
|[<span data-ttu-id="97607-835">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-835">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-836">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="97607-836">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-837">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-837">Example</span></span>

<span data-ttu-id="97607-p154">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="97607-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="97607-841">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="97607-841">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="97607-842">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="97607-842">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="97607-p155">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="97607-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97607-847">Paramètres</span><span class="sxs-lookup"><span data-stu-id="97607-847">Parameters</span></span>

|<span data-ttu-id="97607-848">Nom</span><span class="sxs-lookup"><span data-stu-id="97607-848">Name</span></span>| <span data-ttu-id="97607-849">Type</span><span class="sxs-lookup"><span data-stu-id="97607-849">Type</span></span>| <span data-ttu-id="97607-850">Attributs</span><span class="sxs-lookup"><span data-stu-id="97607-850">Attributes</span></span>| <span data-ttu-id="97607-851">Description</span><span class="sxs-lookup"><span data-stu-id="97607-851">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="97607-852">String</span><span class="sxs-lookup"><span data-stu-id="97607-852">String</span></span>||<span data-ttu-id="97607-853">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="97607-853">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="97607-854">Objet</span><span class="sxs-lookup"><span data-stu-id="97607-854">Object</span></span>| <span data-ttu-id="97607-855">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-855">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-856">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="97607-856">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97607-857">Objet</span><span class="sxs-lookup"><span data-stu-id="97607-857">Object</span></span>| <span data-ttu-id="97607-858">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-858">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-859">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="97607-859">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97607-860">fonction</span><span class="sxs-lookup"><span data-stu-id="97607-860">function</span></span>| <span data-ttu-id="97607-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97607-861">&lt;optional&gt;</span></span>|<span data-ttu-id="97607-862">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="97607-862">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="97607-863">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="97607-863">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="97607-864">Erreurs</span><span class="sxs-lookup"><span data-stu-id="97607-864">Errors</span></span>

| <span data-ttu-id="97607-865">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="97607-865">Error code</span></span> | <span data-ttu-id="97607-866">Description</span><span class="sxs-lookup"><span data-stu-id="97607-866">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="97607-867">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="97607-867">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97607-868">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="97607-868">Requirements</span></span>

|<span data-ttu-id="97607-869">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="97607-869">Requirement</span></span>| <span data-ttu-id="97607-870">Valeur</span><span class="sxs-lookup"><span data-stu-id="97607-870">Value</span></span>|
|---|---|
|[<span data-ttu-id="97607-871">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="97607-871">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97607-872">1.1</span><span class="sxs-lookup"><span data-stu-id="97607-872">1.1</span></span>|
|[<span data-ttu-id="97607-873">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="97607-873">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97607-874">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97607-874">ReadWriteItem</span></span>|
|[<span data-ttu-id="97607-875">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="97607-875">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97607-876">Composition</span><span class="sxs-lookup"><span data-stu-id="97607-876">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97607-877">Exemple</span><span class="sxs-lookup"><span data-stu-id="97607-877">Example</span></span>

<span data-ttu-id="97607-878">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="97607-878">The following code removes an attachment with an identifier of '0'.</span></span>

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
