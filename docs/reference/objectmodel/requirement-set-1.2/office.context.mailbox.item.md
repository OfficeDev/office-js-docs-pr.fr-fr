---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8e411ac1ce58dd59ad3bfc6590a310289bbe686d
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870512"
---
# <a name="item"></a><span data-ttu-id="7f945-102">élément</span><span class="sxs-lookup"><span data-stu-id="7f945-102">item</span></span>

### <span data-ttu-id="7f945-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="7f945-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="7f945-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="7f945-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-107">Requirements</span></span>

|<span data-ttu-id="7f945-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-108">Requirement</span></span>| <span data-ttu-id="7f945-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-111">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-111">1.0</span></span>|
|[<span data-ttu-id="7f945-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7f945-113">Restricted</span></span>|
|[<span data-ttu-id="7f945-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-115">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="7f945-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-116">Example</span></span>

<span data-ttu-id="7f945-117">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="7f945-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="7f945-118">Membres</span><span class="sxs-lookup"><span data-stu-id="7f945-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="7f945-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="7f945-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="7f945-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-122">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="7f945-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="7f945-123">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="7f945-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-124">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-124">Type</span></span>

*   <span data-ttu-id="7f945-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="7f945-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-126">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-126">Requirements</span></span>

|<span data-ttu-id="7f945-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-127">Requirement</span></span>| <span data-ttu-id="7f945-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-129">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-130">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-130">1.0</span></span>|
|[<span data-ttu-id="7f945-131">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-132">ReadItem</span></span>|
|[<span data-ttu-id="7f945-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-134">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-135">Example</span></span>

<span data-ttu-id="7f945-136">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="7f945-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="7f945-138">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="7f945-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="7f945-139">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-140">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-140">Type</span></span>

*   [<span data-ttu-id="7f945-141">Destinataires</span><span class="sxs-lookup"><span data-stu-id="7f945-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="7f945-142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-142">Requirements</span></span>

|<span data-ttu-id="7f945-143">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-143">Requirement</span></span>| <span data-ttu-id="7f945-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-145">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-146">1.1</span><span class="sxs-lookup"><span data-stu-id="7f945-146">1.1</span></span>|
|[<span data-ttu-id="7f945-147">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-148">ReadItem</span></span>|
|[<span data-ttu-id="7f945-149">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-150">Composition</span><span class="sxs-lookup"><span data-stu-id="7f945-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-151">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="7f945-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="7f945-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="7f945-153">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-154">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-154">Type</span></span>

*   [<span data-ttu-id="7f945-155">Body</span><span class="sxs-lookup"><span data-stu-id="7f945-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="7f945-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-156">Requirements</span></span>

|<span data-ttu-id="7f945-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-157">Requirement</span></span>| <span data-ttu-id="7f945-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-160">1.1</span><span class="sxs-lookup"><span data-stu-id="7f945-160">1.1</span></span>|
|[<span data-ttu-id="7f945-161">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-162">ReadItem</span></span>|
|[<span data-ttu-id="7f945-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-165">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-165">Example</span></span>

<span data-ttu-id="7f945-166">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="7f945-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="7f945-167">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="7f945-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="7f945-169">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="7f945-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="7f945-170">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-171">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-171">Read mode</span></span>

<span data-ttu-id="7f945-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7f945-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-174">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-174">Compose mode</span></span>

<span data-ttu-id="7f945-175">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="7f945-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7f945-176">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-176">Type</span></span>

*   <span data-ttu-id="7f945-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-178">Requirements</span></span>

|<span data-ttu-id="7f945-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-179">Requirement</span></span>| <span data-ttu-id="7f945-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-182">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-182">1.0</span></span>|
|[<span data-ttu-id="7f945-183">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-184">ReadItem</span></span>|
|[<span data-ttu-id="7f945-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="7f945-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="7f945-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="7f945-188">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="7f945-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="7f945-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="7f945-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="7f945-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="7f945-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-193">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-193">Type</span></span>

*   <span data-ttu-id="7f945-194">String</span><span class="sxs-lookup"><span data-stu-id="7f945-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-195">Requirements</span></span>

|<span data-ttu-id="7f945-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-196">Requirement</span></span>| <span data-ttu-id="7f945-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-199">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-199">1.0</span></span>|
|[<span data-ttu-id="7f945-200">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-201">ReadItem</span></span>|
|[<span data-ttu-id="7f945-202">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-203">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-204">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="7f945-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="7f945-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="7f945-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-208">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-208">Type</span></span>

*   <span data-ttu-id="7f945-209">Date</span><span class="sxs-lookup"><span data-stu-id="7f945-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-210">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-210">Requirements</span></span>

|<span data-ttu-id="7f945-211">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-211">Requirement</span></span>| <span data-ttu-id="7f945-212">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-213">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-214">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-214">1.0</span></span>|
|[<span data-ttu-id="7f945-215">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-216">ReadItem</span></span>|
|[<span data-ttu-id="7f945-217">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-218">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-219">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="7f945-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="7f945-220">dateTimeModified :Date</span></span>

<span data-ttu-id="7f945-p111">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-223">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-224">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-224">Type</span></span>

*   <span data-ttu-id="7f945-225">Date</span><span class="sxs-lookup"><span data-stu-id="7f945-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-226">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-226">Requirements</span></span>

|<span data-ttu-id="7f945-227">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-227">Requirement</span></span>| <span data-ttu-id="7f945-228">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-229">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-230">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-230">1.0</span></span>|
|[<span data-ttu-id="7f945-231">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-232">ReadItem</span></span>|
|[<span data-ttu-id="7f945-233">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-234">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-235">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="7f945-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="7f945-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="7f945-237">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="7f945-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="7f945-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-240">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-240">Read mode</span></span>

<span data-ttu-id="7f945-241">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="7f945-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-242">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-242">Compose mode</span></span>

<span data-ttu-id="7f945-243">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7f945-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="7f945-244">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="7f945-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="7f945-245">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7f945-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7f945-246">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-246">Type</span></span>

*   <span data-ttu-id="7f945-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="7f945-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-248">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-248">Requirements</span></span>

|<span data-ttu-id="7f945-249">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-249">Requirement</span></span>| <span data-ttu-id="7f945-250">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-251">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-252">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-252">1.0</span></span>|
|[<span data-ttu-id="7f945-253">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-254">ReadItem</span></span>|
|[<span data-ttu-id="7f945-255">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-256">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="7f945-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="7f945-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="7f945-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="7f945-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="7f945-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-262">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7f945-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-263">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-263">Type</span></span>

*   [<span data-ttu-id="7f945-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7f945-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="7f945-265">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-265">Requirements</span></span>

|<span data-ttu-id="7f945-266">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-266">Requirement</span></span>| <span data-ttu-id="7f945-267">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-268">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-269">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-269">1.0</span></span>|
|[<span data-ttu-id="7f945-270">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-271">ReadItem</span></span>|
|[<span data-ttu-id="7f945-272">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-273">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-274">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="7f945-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="7f945-275">internetMessageId :String</span></span>

<span data-ttu-id="7f945-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-278">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-278">Type</span></span>

*   <span data-ttu-id="7f945-279">String</span><span class="sxs-lookup"><span data-stu-id="7f945-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-280">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-280">Requirements</span></span>

|<span data-ttu-id="7f945-281">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-281">Requirement</span></span>| <span data-ttu-id="7f945-282">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-283">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-284">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-284">1.0</span></span>|
|[<span data-ttu-id="7f945-285">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-286">ReadItem</span></span>|
|[<span data-ttu-id="7f945-287">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-288">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-289">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="7f945-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="7f945-290">itemClass :String</span></span>

<span data-ttu-id="7f945-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="7f945-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="7f945-295">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-295">Type</span></span> | <span data-ttu-id="7f945-296">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-296">Description</span></span> | <span data-ttu-id="7f945-297">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="7f945-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="7f945-298">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="7f945-298">Appointment items</span></span> | <span data-ttu-id="7f945-299">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="7f945-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="7f945-300">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="7f945-300">Message items</span></span> | <span data-ttu-id="7f945-301">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="7f945-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="7f945-302">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="7f945-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-303">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-303">Type</span></span>

*   <span data-ttu-id="7f945-304">String</span><span class="sxs-lookup"><span data-stu-id="7f945-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-305">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-305">Requirements</span></span>

|<span data-ttu-id="7f945-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-306">Requirement</span></span>| <span data-ttu-id="7f945-307">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-308">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-309">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-309">1.0</span></span>|
|[<span data-ttu-id="7f945-310">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-311">ReadItem</span></span>|
|[<span data-ttu-id="7f945-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-313">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-314">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="7f945-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="7f945-315">(nullable) itemId :String</span></span>

<span data-ttu-id="7f945-p118">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-318">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="7f945-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="7f945-319">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="7f945-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="7f945-320">Avant d'effectuer des appels d'API REST à l'aide de cette valeur `Office.context.mailbox.convertToRestId`, elle doit être convertie à l'aide de, qui est disponible à partir de l'ensemble de conditions requises 1,3.</span><span class="sxs-lookup"><span data-stu-id="7f945-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="7f945-321">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="7f945-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-322">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-322">Type</span></span>

*   <span data-ttu-id="7f945-323">String</span><span class="sxs-lookup"><span data-stu-id="7f945-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-324">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-324">Requirements</span></span>

|<span data-ttu-id="7f945-325">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-325">Requirement</span></span>| <span data-ttu-id="7f945-326">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-327">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-328">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-328">1.0</span></span>|
|[<span data-ttu-id="7f945-329">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-330">ReadItem</span></span>|
|[<span data-ttu-id="7f945-331">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-332">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-333">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-333">Example</span></span>

<span data-ttu-id="7f945-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7f945-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="7f945-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="7f945-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="7f945-337">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="7f945-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="7f945-338">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-339">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-339">Type</span></span>

*   [<span data-ttu-id="7f945-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="7f945-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="7f945-341">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-341">Requirements</span></span>

|<span data-ttu-id="7f945-342">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-342">Requirement</span></span>| <span data-ttu-id="7f945-343">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-344">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-345">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-345">1.0</span></span>|
|[<span data-ttu-id="7f945-346">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-347">ReadItem</span></span>|
|[<span data-ttu-id="7f945-348">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-349">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-350">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="7f945-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="7f945-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="7f945-352">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-353">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-353">Read mode</span></span>

<span data-ttu-id="7f945-354">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-355">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-355">Compose mode</span></span>

<span data-ttu-id="7f945-356">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7f945-357">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-357">Type</span></span>

*   <span data-ttu-id="7f945-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="7f945-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-359">Requirements</span></span>

|<span data-ttu-id="7f945-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-360">Requirement</span></span>| <span data-ttu-id="7f945-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-363">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-363">1.0</span></span>|
|[<span data-ttu-id="7f945-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-365">ReadItem</span></span>|
|[<span data-ttu-id="7f945-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-367">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="7f945-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="7f945-368">normalizedSubject :String</span></span>

<span data-ttu-id="7f945-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="7f945-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="7f945-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-373">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-373">Type</span></span>

*   <span data-ttu-id="7f945-374">String</span><span class="sxs-lookup"><span data-stu-id="7f945-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-375">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-375">Requirements</span></span>

|<span data-ttu-id="7f945-376">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-376">Requirement</span></span>| <span data-ttu-id="7f945-377">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-378">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-379">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-379">1.0</span></span>|
|[<span data-ttu-id="7f945-380">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-381">ReadItem</span></span>|
|[<span data-ttu-id="7f945-382">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-383">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-384">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="7f945-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="7f945-386">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="7f945-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="7f945-387">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-388">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-388">Read mode</span></span>

<span data-ttu-id="7f945-389">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="7f945-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-390">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-390">Compose mode</span></span>

<span data-ttu-id="7f945-391">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="7f945-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7f945-392">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-392">Type</span></span>

*   <span data-ttu-id="7f945-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-394">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-394">Requirements</span></span>

|<span data-ttu-id="7f945-395">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-395">Requirement</span></span>| <span data-ttu-id="7f945-396">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-397">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-398">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-398">1.0</span></span>|
|[<span data-ttu-id="7f945-399">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-400">ReadItem</span></span>|
|[<span data-ttu-id="7f945-401">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-402">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="7f945-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="7f945-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="7f945-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-406">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-406">Type</span></span>

*   [<span data-ttu-id="7f945-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7f945-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="7f945-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-408">Requirements</span></span>

|<span data-ttu-id="7f945-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-409">Requirement</span></span>| <span data-ttu-id="7f945-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-412">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-412">1.0</span></span>|
|[<span data-ttu-id="7f945-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-414">ReadItem</span></span>|
|[<span data-ttu-id="7f945-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="7f945-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="7f945-419">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="7f945-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="7f945-420">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-421">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-421">Read mode</span></span>

<span data-ttu-id="7f945-422">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="7f945-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-423">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-423">Compose mode</span></span>

<span data-ttu-id="7f945-424">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="7f945-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="7f945-425">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-425">Type</span></span>

*   <span data-ttu-id="7f945-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-427">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-427">Requirements</span></span>

|<span data-ttu-id="7f945-428">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-428">Requirement</span></span>| <span data-ttu-id="7f945-429">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-430">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-431">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-431">1.0</span></span>|
|[<span data-ttu-id="7f945-432">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-433">ReadItem</span></span>|
|[<span data-ttu-id="7f945-434">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-435">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="7f945-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="7f945-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="7f945-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7f945-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="7f945-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="7f945-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-441">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7f945-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7f945-442">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-442">Type</span></span>

*   [<span data-ttu-id="7f945-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7f945-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="7f945-444">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-444">Requirements</span></span>

|<span data-ttu-id="7f945-445">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-445">Requirement</span></span>| <span data-ttu-id="7f945-446">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-447">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-448">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-448">1.0</span></span>|
|[<span data-ttu-id="7f945-449">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-450">ReadItem</span></span>|
|[<span data-ttu-id="7f945-451">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-452">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-453">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="7f945-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="7f945-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="7f945-455">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="7f945-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="7f945-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-458">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-458">Read mode</span></span>

<span data-ttu-id="7f945-459">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="7f945-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-460">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-460">Compose mode</span></span>

<span data-ttu-id="7f945-461">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7f945-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="7f945-462">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="7f945-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="7f945-463">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7f945-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7f945-464">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-464">Type</span></span>

*   <span data-ttu-id="7f945-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="7f945-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-466">Requirements</span></span>

|<span data-ttu-id="7f945-467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-467">Requirement</span></span>| <span data-ttu-id="7f945-468">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-470">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-470">1.0</span></span>|
|[<span data-ttu-id="7f945-471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-472">ReadItem</span></span>|
|[<span data-ttu-id="7f945-473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-474">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="7f945-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="7f945-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="7f945-476">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="7f945-477">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="7f945-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-478">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-478">Read mode</span></span>

<span data-ttu-id="7f945-p130">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="7f945-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-481">Compose mode</span></span>

<span data-ttu-id="7f945-482">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="7f945-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="7f945-483">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-483">Type</span></span>

*   <span data-ttu-id="7f945-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="7f945-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-485">Requirements</span></span>

|<span data-ttu-id="7f945-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-486">Requirement</span></span>| <span data-ttu-id="7f945-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-489">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-489">1.0</span></span>|
|[<span data-ttu-id="7f945-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-491">ReadItem</span></span>|
|[<span data-ttu-id="7f945-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-493">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="7f945-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="7f945-495">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="7f945-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="7f945-496">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7f945-497">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-497">Read mode</span></span>

<span data-ttu-id="7f945-p132">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7f945-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="7f945-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7f945-500">Compose mode</span></span>

<span data-ttu-id="7f945-501">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="7f945-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7f945-502">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-502">Type</span></span>

*   <span data-ttu-id="7f945-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7f945-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-504">Requirements</span></span>

|<span data-ttu-id="7f945-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-505">Requirement</span></span>| <span data-ttu-id="7f945-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-508">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-508">1.0</span></span>|
|[<span data-ttu-id="7f945-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-510">ReadItem</span></span>|
|[<span data-ttu-id="7f945-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-512">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="7f945-513">Méthodes</span><span class="sxs-lookup"><span data-stu-id="7f945-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="7f945-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7f945-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7f945-515">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7f945-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="7f945-516">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="7f945-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="7f945-517">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7f945-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-518">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-518">Parameters</span></span>

|<span data-ttu-id="7f945-519">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-519">Name</span></span>| <span data-ttu-id="7f945-520">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-520">Type</span></span>| <span data-ttu-id="7f945-521">Attributs</span><span class="sxs-lookup"><span data-stu-id="7f945-521">Attributes</span></span>| <span data-ttu-id="7f945-522">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="7f945-523">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-523">String</span></span>||<span data-ttu-id="7f945-p133">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7f945-526">String</span><span class="sxs-lookup"><span data-stu-id="7f945-526">String</span></span>||<span data-ttu-id="7f945-p134">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7f945-529">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-529">Object</span></span>| <span data-ttu-id="7f945-530">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-530">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-531">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7f945-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7f945-532">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-532">Object</span></span>| <span data-ttu-id="7f945-533">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-533">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-534">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7f945-535">fonction</span><span class="sxs-lookup"><span data-stu-id="7f945-535">function</span></span>| <span data-ttu-id="7f945-536">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-536">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-537">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7f945-538">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7f945-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7f945-539">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="7f945-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7f945-540">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7f945-540">Errors</span></span>

| <span data-ttu-id="7f945-541">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="7f945-541">Error code</span></span> | <span data-ttu-id="7f945-542">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="7f945-543">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="7f945-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="7f945-544">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="7f945-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7f945-545">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7f945-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f945-546">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-546">Requirements</span></span>

|<span data-ttu-id="7f945-547">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-547">Requirement</span></span>| <span data-ttu-id="7f945-548">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-549">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-550">1.1</span><span class="sxs-lookup"><span data-stu-id="7f945-550">1.1</span></span>|
|[<span data-ttu-id="7f945-551">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7f945-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="7f945-553">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-554">Composition</span><span class="sxs-lookup"><span data-stu-id="7f945-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-555">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="7f945-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7f945-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7f945-557">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="7f945-p135">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="7f945-561">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7f945-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="7f945-562">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="7f945-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-563">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-563">Parameters</span></span>

|<span data-ttu-id="7f945-564">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-564">Name</span></span>| <span data-ttu-id="7f945-565">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-565">Type</span></span>| <span data-ttu-id="7f945-566">Attributs</span><span class="sxs-lookup"><span data-stu-id="7f945-566">Attributes</span></span>| <span data-ttu-id="7f945-567">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="7f945-568">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-568">String</span></span>||<span data-ttu-id="7f945-p136">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7f945-571">String</span><span class="sxs-lookup"><span data-stu-id="7f945-571">String</span></span>||<span data-ttu-id="7f945-572">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="7f945-572">The subject of the item to be attached.</span></span> <span data-ttu-id="7f945-573">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7f945-574">Object</span><span class="sxs-lookup"><span data-stu-id="7f945-574">Object</span></span>| <span data-ttu-id="7f945-575">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-575">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-576">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7f945-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7f945-577">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-577">Object</span></span>| <span data-ttu-id="7f945-578">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-578">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-579">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7f945-580">fonction</span><span class="sxs-lookup"><span data-stu-id="7f945-580">function</span></span>| <span data-ttu-id="7f945-581">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-581">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-582">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7f945-583">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7f945-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7f945-584">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="7f945-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7f945-585">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7f945-585">Errors</span></span>

| <span data-ttu-id="7f945-586">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="7f945-586">Error code</span></span> | <span data-ttu-id="7f945-587">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7f945-588">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7f945-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f945-589">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-589">Requirements</span></span>

|<span data-ttu-id="7f945-590">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-590">Requirement</span></span>| <span data-ttu-id="7f945-591">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-592">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-593">1.1</span><span class="sxs-lookup"><span data-stu-id="7f945-593">1.1</span></span>|
|[<span data-ttu-id="7f945-594">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7f945-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="7f945-596">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-597">Composition</span><span class="sxs-lookup"><span data-stu-id="7f945-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-598">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-598">Example</span></span>

<span data-ttu-id="7f945-599">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="7f945-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="7f945-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7f945-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="7f945-601">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7f945-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-602">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7f945-603">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="7f945-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7f945-604">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="7f945-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="7f945-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="7f945-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-608">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-608">Parameters</span></span>

|<span data-ttu-id="7f945-609">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-609">Name</span></span>| <span data-ttu-id="7f945-610">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-610">Type</span></span>| <span data-ttu-id="7f945-611">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="7f945-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7f945-612">String &#124; Object</span></span>| |<span data-ttu-id="7f945-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7f945-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7f945-615">**OU**</span><span class="sxs-lookup"><span data-stu-id="7f945-615">**OR**</span></span><br/><span data-ttu-id="7f945-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="7f945-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7f945-618">String</span><span class="sxs-lookup"><span data-stu-id="7f945-618">String</span></span> | <span data-ttu-id="7f945-619">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-619">&lt;optional&gt;</span></span> | <span data-ttu-id="7f945-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7f945-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7f945-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7f945-623">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-623">&lt;optional&gt;</span></span> | <span data-ttu-id="7f945-624">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7f945-625">String</span><span class="sxs-lookup"><span data-stu-id="7f945-625">String</span></span> | | <span data-ttu-id="7f945-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7f945-628">String</span><span class="sxs-lookup"><span data-stu-id="7f945-628">String</span></span> | | <span data-ttu-id="7f945-629">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7f945-630">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-630">String</span></span> | | <span data-ttu-id="7f945-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="7f945-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7f945-633">String</span><span class="sxs-lookup"><span data-stu-id="7f945-633">String</span></span> | | <span data-ttu-id="7f945-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7f945-637">function</span><span class="sxs-lookup"><span data-stu-id="7f945-637">function</span></span> | <span data-ttu-id="7f945-638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-638">&lt;optional&gt;</span></span> | <span data-ttu-id="7f945-639">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f945-640">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-640">Requirements</span></span>

|<span data-ttu-id="7f945-641">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-641">Requirement</span></span>| <span data-ttu-id="7f945-642">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-643">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-644">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-644">1.0</span></span>|
|[<span data-ttu-id="7f945-645">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-646">ReadItem</span></span>|
|[<span data-ttu-id="7f945-647">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-648">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7f945-649">Exemples</span><span class="sxs-lookup"><span data-stu-id="7f945-649">Examples</span></span>

<span data-ttu-id="7f945-650">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="7f945-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="7f945-651">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="7f945-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="7f945-652">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="7f945-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7f945-653">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="7f945-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7f945-654">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7f945-655">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="7f945-656">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7f945-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="7f945-657">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7f945-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-658">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-658">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7f945-659">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="7f945-659">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7f945-660">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="7f945-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="7f945-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="7f945-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-664">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-664">Parameters</span></span>

|<span data-ttu-id="7f945-665">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-665">Name</span></span>| <span data-ttu-id="7f945-666">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-666">Type</span></span>| <span data-ttu-id="7f945-667">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="7f945-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7f945-668">String &#124; Object</span></span>| | <span data-ttu-id="7f945-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7f945-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7f945-671">**OU**</span><span class="sxs-lookup"><span data-stu-id="7f945-671">**OR**</span></span><br/><span data-ttu-id="7f945-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="7f945-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7f945-674">String</span><span class="sxs-lookup"><span data-stu-id="7f945-674">String</span></span> | <span data-ttu-id="7f945-675">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-675">&lt;optional&gt;</span></span> | <span data-ttu-id="7f945-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7f945-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7f945-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7f945-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-679">&lt;optional&gt;</span></span> | <span data-ttu-id="7f945-680">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7f945-681">String</span><span class="sxs-lookup"><span data-stu-id="7f945-681">String</span></span> | | <span data-ttu-id="7f945-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7f945-684">String</span><span class="sxs-lookup"><span data-stu-id="7f945-684">String</span></span> | | <span data-ttu-id="7f945-685">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7f945-686">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-686">String</span></span> | | <span data-ttu-id="7f945-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="7f945-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7f945-689">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-689">String</span></span> | | <span data-ttu-id="7f945-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7f945-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7f945-693">function</span><span class="sxs-lookup"><span data-stu-id="7f945-693">function</span></span> | <span data-ttu-id="7f945-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-694">&lt;optional&gt;</span></span> | <span data-ttu-id="7f945-695">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f945-696">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-696">Requirements</span></span>

|<span data-ttu-id="7f945-697">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-697">Requirement</span></span>| <span data-ttu-id="7f945-698">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-699">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-700">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-700">1.0</span></span>|
|[<span data-ttu-id="7f945-701">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-702">ReadItem</span></span>|
|[<span data-ttu-id="7f945-703">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-704">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7f945-705">Exemples</span><span class="sxs-lookup"><span data-stu-id="7f945-705">Examples</span></span>

<span data-ttu-id="7f945-706">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="7f945-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="7f945-707">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="7f945-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="7f945-708">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="7f945-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7f945-709">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="7f945-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7f945-710">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7f945-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7f945-711">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="7f945-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="7f945-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="7f945-713">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7f945-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-714">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-714">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-715">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-715">Requirements</span></span>

|<span data-ttu-id="7f945-716">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-716">Requirement</span></span>| <span data-ttu-id="7f945-717">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-718">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-719">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-719">1.0</span></span>|
|[<span data-ttu-id="7f945-720">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-720">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-721">ReadItem</span></span>|
|[<span data-ttu-id="7f945-722">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-722">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-723">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f945-724">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7f945-724">Returns:</span></span>

<span data-ttu-id="7f945-725">Type : [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="7f945-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="7f945-726">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-726">Example</span></span>

<span data-ttu-id="7f945-727">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="7f945-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="7f945-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="7f945-729">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7f945-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-730">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-730">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-731">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-731">Parameters</span></span>

|<span data-ttu-id="7f945-732">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-732">Name</span></span>| <span data-ttu-id="7f945-733">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-733">Type</span></span>| <span data-ttu-id="7f945-734">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="7f945-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="7f945-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="7f945-736">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="7f945-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f945-737">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-737">Requirements</span></span>

|<span data-ttu-id="7f945-738">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-738">Requirement</span></span>| <span data-ttu-id="7f945-739">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-740">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-741">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-741">1.0</span></span>|
|[<span data-ttu-id="7f945-742">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-743">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7f945-743">Restricted</span></span>|
|[<span data-ttu-id="7f945-744">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-745">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f945-746">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7f945-746">Returns:</span></span>

<span data-ttu-id="7f945-747">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="7f945-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="7f945-748">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="7f945-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="7f945-749">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="7f945-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="7f945-750">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="7f945-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="7f945-751">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="7f945-751">Value of `entityType`</span></span> | <span data-ttu-id="7f945-752">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="7f945-752">Type of objects in returned array</span></span> | <span data-ttu-id="7f945-753">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="7f945-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="7f945-754">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-754">String</span></span> | <span data-ttu-id="7f945-755">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7f945-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="7f945-756">Contact</span><span class="sxs-lookup"><span data-stu-id="7f945-756">Contact</span></span> | <span data-ttu-id="7f945-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7f945-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="7f945-758">String</span><span class="sxs-lookup"><span data-stu-id="7f945-758">String</span></span> | <span data-ttu-id="7f945-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7f945-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="7f945-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="7f945-760">MeetingSuggestion</span></span> | <span data-ttu-id="7f945-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7f945-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="7f945-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="7f945-762">PhoneNumber</span></span> | <span data-ttu-id="7f945-763">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7f945-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="7f945-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="7f945-764">TaskSuggestion</span></span> | <span data-ttu-id="7f945-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7f945-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="7f945-766">String</span><span class="sxs-lookup"><span data-stu-id="7f945-766">String</span></span> | <span data-ttu-id="7f945-767">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7f945-767">**Restricted**</span></span> |

<span data-ttu-id="7f945-768">Type : Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="7f945-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="7f945-769">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-769">Example</span></span>

<span data-ttu-id="7f945-770">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7f945-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="7f945-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="7f945-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="7f945-772">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7f945-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-773">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-773">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7f945-774">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="7f945-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-775">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-775">Parameters</span></span>

|<span data-ttu-id="7f945-776">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-776">Name</span></span>| <span data-ttu-id="7f945-777">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-777">Type</span></span>| <span data-ttu-id="7f945-778">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7f945-779">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-779">String</span></span>|<span data-ttu-id="7f945-780">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="7f945-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f945-781">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-781">Requirements</span></span>

|<span data-ttu-id="7f945-782">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-782">Requirement</span></span>| <span data-ttu-id="7f945-783">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-784">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-785">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-785">1.0</span></span>|
|[<span data-ttu-id="7f945-786">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-786">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-787">ReadItem</span></span>|
|[<span data-ttu-id="7f945-788">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-788">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-789">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f945-790">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7f945-790">Returns:</span></span>

<span data-ttu-id="7f945-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="7f945-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="7f945-793">Type : Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="7f945-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="7f945-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="7f945-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="7f945-795">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7f945-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-796">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-796">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7f945-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="7f945-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="7f945-800">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="7f945-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="7f945-801">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="7f945-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="7f945-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="7f945-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f945-804">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-804">Requirements</span></span>

|<span data-ttu-id="7f945-805">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-805">Requirement</span></span>| <span data-ttu-id="7f945-806">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-807">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-808">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-808">1.0</span></span>|
|[<span data-ttu-id="7f945-809">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-810">ReadItem</span></span>|
|[<span data-ttu-id="7f945-811">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-812">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f945-813">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7f945-813">Returns:</span></span>

<span data-ttu-id="7f945-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="7f945-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="7f945-816">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7f945-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7f945-817">Object</span><span class="sxs-lookup"><span data-stu-id="7f945-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="7f945-818">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-818">Example</span></span>

<span data-ttu-id="7f945-819">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="7f945-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="7f945-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="7f945-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="7f945-821">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7f945-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7f945-822">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7f945-822">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7f945-823">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="7f945-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="7f945-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="7f945-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-826">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-826">Parameters</span></span>

|<span data-ttu-id="7f945-827">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-827">Name</span></span>| <span data-ttu-id="7f945-828">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-828">Type</span></span>| <span data-ttu-id="7f945-829">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7f945-830">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-830">String</span></span>|<span data-ttu-id="7f945-831">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="7f945-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f945-832">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-832">Requirements</span></span>

|<span data-ttu-id="7f945-833">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-833">Requirement</span></span>| <span data-ttu-id="7f945-834">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-835">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-836">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-836">1.0</span></span>|
|[<span data-ttu-id="7f945-837">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-837">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-838">ReadItem</span></span>|
|[<span data-ttu-id="7f945-839">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-839">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-840">Lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f945-841">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7f945-841">Returns:</span></span>

<span data-ttu-id="7f945-842">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7f945-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="7f945-843">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7f945-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7f945-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="7f945-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="7f945-845">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="7f945-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="7f945-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="7f945-847">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="7f945-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="7f945-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="7f945-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-850">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-850">Parameters</span></span>

|<span data-ttu-id="7f945-851">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-851">Name</span></span>| <span data-ttu-id="7f945-852">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-852">Type</span></span>| <span data-ttu-id="7f945-853">Attributs</span><span class="sxs-lookup"><span data-stu-id="7f945-853">Attributes</span></span>| <span data-ttu-id="7f945-854">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="7f945-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7f945-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="7f945-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="7f945-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="7f945-859">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-859">Object</span></span>| <span data-ttu-id="7f945-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-860">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-861">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7f945-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7f945-862">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-862">Object</span></span>| <span data-ttu-id="7f945-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-863">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-864">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7f945-865">fonction</span><span class="sxs-lookup"><span data-stu-id="7f945-865">function</span></span>||<span data-ttu-id="7f945-866">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7f945-867">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="7f945-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="7f945-868">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="7f945-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f945-869">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-869">Requirements</span></span>

|<span data-ttu-id="7f945-870">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-870">Requirement</span></span>| <span data-ttu-id="7f945-871">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-872">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-873">1.2</span><span class="sxs-lookup"><span data-stu-id="7f945-873">1.2</span></span>|
|[<span data-ttu-id="7f945-874">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-874">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7f945-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="7f945-876">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-876">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-877">Composition</span><span class="sxs-lookup"><span data-stu-id="7f945-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="7f945-878">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="7f945-878">Returns:</span></span>

<span data-ttu-id="7f945-879">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="7f945-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="7f945-880">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7f945-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7f945-881">String</span><span class="sxs-lookup"><span data-stu-id="7f945-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="7f945-882">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-882">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="7f945-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7f945-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="7f945-884">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7f945-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="7f945-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="7f945-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-888">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-888">Parameters</span></span>

|<span data-ttu-id="7f945-889">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-889">Name</span></span>| <span data-ttu-id="7f945-890">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-890">Type</span></span>| <span data-ttu-id="7f945-891">Attributs</span><span class="sxs-lookup"><span data-stu-id="7f945-891">Attributes</span></span>| <span data-ttu-id="7f945-892">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7f945-893">function</span><span class="sxs-lookup"><span data-stu-id="7f945-893">function</span></span>||<span data-ttu-id="7f945-894">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7f945-895">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7f945-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="7f945-896">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="7f945-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="7f945-897">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-897">Object</span></span>| <span data-ttu-id="7f945-898">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-898">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-899">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="7f945-900">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f945-901">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-901">Requirements</span></span>

|<span data-ttu-id="7f945-902">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-902">Requirement</span></span>| <span data-ttu-id="7f945-903">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-904">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-905">1.0</span><span class="sxs-lookup"><span data-stu-id="7f945-905">1.0</span></span>|
|[<span data-ttu-id="7f945-906">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-906">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f945-907">ReadItem</span></span>|
|[<span data-ttu-id="7f945-908">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-908">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-909">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7f945-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-910">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-910">Example</span></span>

<span data-ttu-id="7f945-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="7f945-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="7f945-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7f945-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="7f945-915">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7f945-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="7f945-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="7f945-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-920">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-920">Parameters</span></span>

|<span data-ttu-id="7f945-921">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-921">Name</span></span>| <span data-ttu-id="7f945-922">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-922">Type</span></span>| <span data-ttu-id="7f945-923">Attributs</span><span class="sxs-lookup"><span data-stu-id="7f945-923">Attributes</span></span>| <span data-ttu-id="7f945-924">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="7f945-925">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7f945-925">String</span></span>||<span data-ttu-id="7f945-926">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="7f945-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="7f945-927">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-927">Object</span></span>| <span data-ttu-id="7f945-928">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-928">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-929">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7f945-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7f945-930">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-930">Object</span></span>| <span data-ttu-id="7f945-931">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-931">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-932">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7f945-933">fonction</span><span class="sxs-lookup"><span data-stu-id="7f945-933">function</span></span>| <span data-ttu-id="7f945-934">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-934">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-935">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7f945-936">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="7f945-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7f945-937">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7f945-937">Errors</span></span>

| <span data-ttu-id="7f945-938">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="7f945-938">Error code</span></span> | <span data-ttu-id="7f945-939">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="7f945-940">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="7f945-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f945-941">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-941">Requirements</span></span>

|<span data-ttu-id="7f945-942">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-942">Requirement</span></span>| <span data-ttu-id="7f945-943">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-944">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-945">1.1</span><span class="sxs-lookup"><span data-stu-id="7f945-945">1.1</span></span>|
|[<span data-ttu-id="7f945-946">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7f945-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="7f945-948">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-949">Composition</span><span class="sxs-lookup"><span data-stu-id="7f945-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-950">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-950">Example</span></span>

<span data-ttu-id="7f945-951">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="7f945-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="7f945-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="7f945-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="7f945-953">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7f945-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="7f945-p166">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="7f945-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7f945-957">Paramètres</span><span class="sxs-lookup"><span data-stu-id="7f945-957">Parameters</span></span>

|<span data-ttu-id="7f945-958">Nom</span><span class="sxs-lookup"><span data-stu-id="7f945-958">Name</span></span>| <span data-ttu-id="7f945-959">Type</span><span class="sxs-lookup"><span data-stu-id="7f945-959">Type</span></span>| <span data-ttu-id="7f945-960">Attributs</span><span class="sxs-lookup"><span data-stu-id="7f945-960">Attributes</span></span>| <span data-ttu-id="7f945-961">Description</span><span class="sxs-lookup"><span data-stu-id="7f945-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7f945-962">String</span><span class="sxs-lookup"><span data-stu-id="7f945-962">String</span></span>||<span data-ttu-id="7f945-p167">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="7f945-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="7f945-966">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-966">Object</span></span>| <span data-ttu-id="7f945-967">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-967">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-968">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7f945-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7f945-969">Objet</span><span class="sxs-lookup"><span data-stu-id="7f945-969">Object</span></span>| <span data-ttu-id="7f945-970">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-970">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-971">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7f945-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="7f945-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7f945-972">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="7f945-973">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="7f945-973">&lt;optional&gt;</span></span>|<span data-ttu-id="7f945-p168">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="7f945-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="7f945-p169">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="7f945-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="7f945-978">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="7f945-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="7f945-979">fonction</span><span class="sxs-lookup"><span data-stu-id="7f945-979">function</span></span>||<span data-ttu-id="7f945-980">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7f945-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f945-981">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7f945-981">Requirements</span></span>

|<span data-ttu-id="7f945-982">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7f945-982">Requirement</span></span>| <span data-ttu-id="7f945-983">Valeur</span><span class="sxs-lookup"><span data-stu-id="7f945-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f945-984">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7f945-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f945-985">1.2</span><span class="sxs-lookup"><span data-stu-id="7f945-985">1.2</span></span>|
|[<span data-ttu-id="7f945-986">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7f945-986">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f945-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7f945-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="7f945-988">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7f945-988">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f945-989">Composition</span><span class="sxs-lookup"><span data-stu-id="7f945-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7f945-990">Exemple</span><span class="sxs-lookup"><span data-stu-id="7f945-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
