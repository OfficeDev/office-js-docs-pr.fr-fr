---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cd8e71e39940fcf0de50982ef1cdb6825abb7221
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450365"
---
# <a name="item"></a><span data-ttu-id="12c6b-102">élément</span><span class="sxs-lookup"><span data-stu-id="12c6b-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="12c6b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="12c6b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="12c6b-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="12c6b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-106">Requirements</span></span>

|<span data-ttu-id="12c6b-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-107">Requirement</span></span>| <span data-ttu-id="12c6b-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-110">1.0</span></span>|
|[<span data-ttu-id="12c6b-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="12c6b-112">Restricted</span></span>|
|[<span data-ttu-id="12c6b-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="12c6b-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-115">Example</span></span>

<span data-ttu-id="12c6b-116">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="12c6b-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="12c6b-117">Membres</span><span class="sxs-lookup"><span data-stu-id="12c6b-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="12c6b-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="12c6b-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="12c6b-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-121">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="12c6b-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="12c6b-122">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="12c6b-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-123">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-123">Type</span></span>

*   <span data-ttu-id="12c6b-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="12c6b-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-125">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-125">Requirements</span></span>

|<span data-ttu-id="12c6b-126">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-126">Requirement</span></span>| <span data-ttu-id="12c6b-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-128">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-129">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-129">1.0</span></span>|
|[<span data-ttu-id="12c6b-130">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-131">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-134">Example</span></span>

<span data-ttu-id="12c6b-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="12c6b-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="12c6b-137">Obtient un objet qui fournit des méthodes permettant d'obtenir ou de mettre à jour la ligne CCI (copie carbone invisible) d'un message.</span><span class="sxs-lookup"><span data-stu-id="12c6b-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="12c6b-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-139">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-139">Type</span></span>

*   [<span data-ttu-id="12c6b-140">Destinataires</span><span class="sxs-lookup"><span data-stu-id="12c6b-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="12c6b-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-141">Requirements</span></span>

|<span data-ttu-id="12c6b-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-142">Requirement</span></span>| <span data-ttu-id="12c6b-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-145">1.1</span><span class="sxs-lookup"><span data-stu-id="12c6b-145">1.1</span></span>|
|[<span data-ttu-id="12c6b-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-147">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-149">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="12c6b-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="12c6b-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="12c6b-152">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-153">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-153">Type</span></span>

*   [<span data-ttu-id="12c6b-154">Body</span><span class="sxs-lookup"><span data-stu-id="12c6b-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="12c6b-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-155">Requirements</span></span>

|<span data-ttu-id="12c6b-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-156">Requirement</span></span>| <span data-ttu-id="12c6b-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-159">1.1</span><span class="sxs-lookup"><span data-stu-id="12c6b-159">1.1</span></span>|
|[<span data-ttu-id="12c6b-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-161">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-164">Example</span></span>

<span data-ttu-id="12c6b-165">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="12c6b-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="12c6b-166">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="12c6b-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="12c6b-168">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="12c6b-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="12c6b-169">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-170">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-170">Read mode</span></span>

<span data-ttu-id="12c6b-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-173">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-173">Compose mode</span></span>

<span data-ttu-id="12c6b-174">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="12c6b-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="12c6b-175">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-175">Type</span></span>

*   <span data-ttu-id="12c6b-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-177">Requirements</span></span>

|<span data-ttu-id="12c6b-178">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-178">Requirement</span></span>| <span data-ttu-id="12c6b-179">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-180">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-181">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-181">1.0</span></span>|
|[<span data-ttu-id="12c6b-182">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-183">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-185">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="12c6b-186">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="12c6b-186">(nullable) conversationId :String</span></span>

<span data-ttu-id="12c6b-187">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="12c6b-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="12c6b-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="12c6b-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-192">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-192">Type</span></span>

*   <span data-ttu-id="12c6b-193">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-194">Requirements</span></span>

|<span data-ttu-id="12c6b-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-195">Requirement</span></span>| <span data-ttu-id="12c6b-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-198">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-198">1.0</span></span>|
|[<span data-ttu-id="12c6b-199">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-200">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="12c6b-204">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="12c6b-204">dateTimeCreated :Date</span></span>

<span data-ttu-id="12c6b-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-207">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-207">Type</span></span>

*   <span data-ttu-id="12c6b-208">Date</span><span class="sxs-lookup"><span data-stu-id="12c6b-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-209">Requirements</span></span>

|<span data-ttu-id="12c6b-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-210">Requirement</span></span>| <span data-ttu-id="12c6b-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-213">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-213">1.0</span></span>|
|[<span data-ttu-id="12c6b-214">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-215">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-217">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-218">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="12c6b-219">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="12c6b-219">dateTimeModified :Date</span></span>

<span data-ttu-id="12c6b-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-222">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-223">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-223">Type</span></span>

*   <span data-ttu-id="12c6b-224">Date</span><span class="sxs-lookup"><span data-stu-id="12c6b-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-225">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-225">Requirements</span></span>

|<span data-ttu-id="12c6b-226">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-226">Requirement</span></span>| <span data-ttu-id="12c6b-227">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-228">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-229">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-229">1.0</span></span>|
|[<span data-ttu-id="12c6b-230">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-231">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-232">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-233">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-234">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="12c6b-235">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="12c6b-235">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="12c6b-236">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="12c6b-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-239">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-239">Read mode</span></span>

<span data-ttu-id="12c6b-240">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-241">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-241">Compose mode</span></span>

<span data-ttu-id="12c6b-242">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="12c6b-243">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="12c6b-244">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="12c6b-245">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-245">Type</span></span>

*   <span data-ttu-id="12c6b-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="12c6b-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-247">Requirements</span></span>

|<span data-ttu-id="12c6b-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-248">Requirement</span></span>| <span data-ttu-id="12c6b-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-251">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-251">1.0</span></span>|
|[<span data-ttu-id="12c6b-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-253">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="12c6b-256">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="12c6b-256">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="12c6b-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="12c6b-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-261">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-262">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-262">Type</span></span>

*   [<span data-ttu-id="12c6b-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="12c6b-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="12c6b-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-264">Requirements</span></span>

|<span data-ttu-id="12c6b-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-265">Requirement</span></span>| <span data-ttu-id="12c6b-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-268">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-268">1.0</span></span>|
|[<span data-ttu-id="12c6b-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-270">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-272">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-273">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="12c6b-274">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="12c6b-274">internetMessageId :String</span></span>

<span data-ttu-id="12c6b-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-277">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-277">Type</span></span>

*   <span data-ttu-id="12c6b-278">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-279">Requirements</span></span>

|<span data-ttu-id="12c6b-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-280">Requirement</span></span>| <span data-ttu-id="12c6b-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-283">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-283">1.0</span></span>|
|[<span data-ttu-id="12c6b-284">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-285">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-287">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="12c6b-289">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="12c6b-289">itemClass :String</span></span>

<span data-ttu-id="12c6b-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="12c6b-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="12c6b-294">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-294">Type</span></span> | <span data-ttu-id="12c6b-295">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-295">Description</span></span> | <span data-ttu-id="12c6b-296">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="12c6b-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="12c6b-297">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="12c6b-297">Appointment items</span></span> | <span data-ttu-id="12c6b-298">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="12c6b-299">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="12c6b-299">Message items</span></span> | <span data-ttu-id="12c6b-300">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="12c6b-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="12c6b-301">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-302">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-302">Type</span></span>

*   <span data-ttu-id="12c6b-303">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-304">Requirements</span></span>

|<span data-ttu-id="12c6b-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-305">Requirement</span></span>| <span data-ttu-id="12c6b-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-308">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-308">1.0</span></span>|
|[<span data-ttu-id="12c6b-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-310">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="12c6b-314">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="12c6b-314">(nullable) itemId :String</span></span>

<span data-ttu-id="12c6b-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-317">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="12c6b-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="12c6b-318">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="12c6b-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="12c6b-319">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="12c6b-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="12c6b-320">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="12c6b-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="12c6b-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-323">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-323">Type</span></span>

*   <span data-ttu-id="12c6b-324">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-325">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-325">Requirements</span></span>

|<span data-ttu-id="12c6b-326">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-326">Requirement</span></span>| <span data-ttu-id="12c6b-327">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-328">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-329">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-329">1.0</span></span>|
|[<span data-ttu-id="12c6b-330">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-331">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-332">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-333">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-334">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-334">Example</span></span>

<span data-ttu-id="12c6b-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="12c6b-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="12c6b-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="12c6b-338">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="12c6b-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="12c6b-339">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-340">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-340">Type</span></span>

*   [<span data-ttu-id="12c6b-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="12c6b-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="12c6b-342">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-342">Requirements</span></span>

|<span data-ttu-id="12c6b-343">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-343">Requirement</span></span>| <span data-ttu-id="12c6b-344">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-345">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-346">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-346">1.0</span></span>|
|[<span data-ttu-id="12c6b-347">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-348">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-349">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-350">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-351">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="12c6b-352">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="12c6b-352">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="12c6b-353">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-354">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-354">Read mode</span></span>

<span data-ttu-id="12c6b-355">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-356">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-356">Compose mode</span></span>

<span data-ttu-id="12c6b-357">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="12c6b-358">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-358">Type</span></span>

*   <span data-ttu-id="12c6b-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="12c6b-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-360">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-360">Requirements</span></span>

|<span data-ttu-id="12c6b-361">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-361">Requirement</span></span>| <span data-ttu-id="12c6b-362">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-363">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-364">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-364">1.0</span></span>|
|[<span data-ttu-id="12c6b-365">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-366">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-367">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-368">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="12c6b-369">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="12c6b-369">normalizedSubject :String</span></span>

<span data-ttu-id="12c6b-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="12c6b-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="12c6b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-374">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-374">Type</span></span>

*   <span data-ttu-id="12c6b-375">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-376">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-376">Requirements</span></span>

|<span data-ttu-id="12c6b-377">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-377">Requirement</span></span>| <span data-ttu-id="12c6b-378">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-379">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-380">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-380">1.0</span></span>|
|[<span data-ttu-id="12c6b-381">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-382">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-383">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-384">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-385">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="12c6b-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="12c6b-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="12c6b-387">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-388">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-388">Type</span></span>

*   [<span data-ttu-id="12c6b-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="12c6b-389">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="12c6b-390">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-390">Requirements</span></span>

|<span data-ttu-id="12c6b-391">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-391">Requirement</span></span>| <span data-ttu-id="12c6b-392">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-393">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-394">1.3</span><span class="sxs-lookup"><span data-stu-id="12c6b-394">1.3</span></span>|
|[<span data-ttu-id="12c6b-395">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-396">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-397">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-398">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-399">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="12c6b-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="12c6b-401">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="12c6b-402">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-403">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-403">Read mode</span></span>

<span data-ttu-id="12c6b-404">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="12c6b-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-405">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-405">Compose mode</span></span>

<span data-ttu-id="12c6b-406">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="12c6b-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="12c6b-407">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-407">Type</span></span>

*   <span data-ttu-id="12c6b-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-409">Requirements</span></span>

|<span data-ttu-id="12c6b-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-410">Requirement</span></span>| <span data-ttu-id="12c6b-411">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-412">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-413">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-413">1.0</span></span>|
|[<span data-ttu-id="12c6b-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-415">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-416">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-417">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="12c6b-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="12c6b-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="12c6b-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-421">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-421">Type</span></span>

*   [<span data-ttu-id="12c6b-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="12c6b-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="12c6b-423">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-423">Requirements</span></span>

|<span data-ttu-id="12c6b-424">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-424">Requirement</span></span>| <span data-ttu-id="12c6b-425">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-426">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-427">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-427">1.0</span></span>|
|[<span data-ttu-id="12c6b-428">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-429">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-431">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-432">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="12c6b-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="12c6b-434">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="12c6b-435">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-436">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-436">Read mode</span></span>

<span data-ttu-id="12c6b-437">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="12c6b-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-438">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-438">Compose mode</span></span>

<span data-ttu-id="12c6b-439">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="12c6b-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="12c6b-440">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-440">Type</span></span>

*   <span data-ttu-id="12c6b-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-442">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-442">Requirements</span></span>

|<span data-ttu-id="12c6b-443">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-443">Requirement</span></span>| <span data-ttu-id="12c6b-444">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-445">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-446">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-446">1.0</span></span>|
|[<span data-ttu-id="12c6b-447">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-448">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-449">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-450">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="12c6b-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="12c6b-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="12c6b-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="12c6b-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-456">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="12c6b-457">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-457">Type</span></span>

*   [<span data-ttu-id="12c6b-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="12c6b-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="12c6b-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-459">Requirements</span></span>

|<span data-ttu-id="12c6b-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-460">Requirement</span></span>| <span data-ttu-id="12c6b-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-463">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-463">1.0</span></span>|
|[<span data-ttu-id="12c6b-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-465">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-467">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="12c6b-469">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="12c6b-469">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="12c6b-470">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="12c6b-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-473">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-473">Read mode</span></span>

<span data-ttu-id="12c6b-474">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-475">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-475">Compose mode</span></span>

<span data-ttu-id="12c6b-476">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="12c6b-477">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="12c6b-478">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="12c6b-479">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-479">Type</span></span>

*   <span data-ttu-id="12c6b-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="12c6b-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-481">Requirements</span></span>

|<span data-ttu-id="12c6b-482">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-482">Requirement</span></span>| <span data-ttu-id="12c6b-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-485">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-485">1.0</span></span>|
|[<span data-ttu-id="12c6b-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-487">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-489">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-489">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="12c6b-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="12c6b-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="12c6b-491">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="12c6b-492">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="12c6b-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-493">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-493">Read mode</span></span>

<span data-ttu-id="12c6b-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-496">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-496">Compose mode</span></span>

<span data-ttu-id="12c6b-497">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="12c6b-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="12c6b-498">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-498">Type</span></span>

*   <span data-ttu-id="12c6b-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="12c6b-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-500">Requirements</span></span>

|<span data-ttu-id="12c6b-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-501">Requirement</span></span>| <span data-ttu-id="12c6b-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-504">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-504">1.0</span></span>|
|[<span data-ttu-id="12c6b-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-506">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-508">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-508">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="12c6b-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="12c6b-510">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="12c6b-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="12c6b-511">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="12c6b-512">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-512">Read mode</span></span>

<span data-ttu-id="12c6b-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="12c6b-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-515">Compose mode</span></span>

<span data-ttu-id="12c6b-516">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="12c6b-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="12c6b-517">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-517">Type</span></span>

*   <span data-ttu-id="12c6b-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="12c6b-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-519">Requirements</span></span>

|<span data-ttu-id="12c6b-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-520">Requirement</span></span>| <span data-ttu-id="12c6b-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-523">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-523">1.0</span></span>|
|[<span data-ttu-id="12c6b-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-525">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-527">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="12c6b-528">Méthodes</span><span class="sxs-lookup"><span data-stu-id="12c6b-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="12c6b-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="12c6b-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="12c6b-530">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="12c6b-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="12c6b-531">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="12c6b-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="12c6b-532">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="12c6b-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-533">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-533">Parameters</span></span>

|<span data-ttu-id="12c6b-534">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-534">Name</span></span>| <span data-ttu-id="12c6b-535">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-535">Type</span></span>| <span data-ttu-id="12c6b-536">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-536">Attributes</span></span>| <span data-ttu-id="12c6b-537">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="12c6b-538">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-538">String</span></span>||<span data-ttu-id="12c6b-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="12c6b-541">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-541">String</span></span>||<span data-ttu-id="12c6b-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="12c6b-544">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-544">Object</span></span>| <span data-ttu-id="12c6b-545">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-545">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-546">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="12c6b-547">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-547">Object</span></span>| <span data-ttu-id="12c6b-548">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-548">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-549">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="12c6b-550">fonction</span><span class="sxs-lookup"><span data-stu-id="12c6b-550">function</span></span>| <span data-ttu-id="12c6b-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-551">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-552">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="12c6b-553">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="12c6b-554">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="12c6b-555">Erreurs</span><span class="sxs-lookup"><span data-stu-id="12c6b-555">Errors</span></span>

| <span data-ttu-id="12c6b-556">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="12c6b-556">Error code</span></span> | <span data-ttu-id="12c6b-557">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="12c6b-558">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="12c6b-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="12c6b-559">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="12c6b-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="12c6b-560">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="12c6b-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-561">Requirements</span></span>

|<span data-ttu-id="12c6b-562">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-562">Requirement</span></span>| <span data-ttu-id="12c6b-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-565">1.1</span><span class="sxs-lookup"><span data-stu-id="12c6b-565">1.1</span></span>|
|[<span data-ttu-id="12c6b-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="12c6b-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-569">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="12c6b-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="12c6b-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="12c6b-572">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="12c6b-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="12c6b-576">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="12c6b-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="12c6b-577">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="12c6b-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-578">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-578">Parameters</span></span>

|<span data-ttu-id="12c6b-579">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-579">Name</span></span>| <span data-ttu-id="12c6b-580">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-580">Type</span></span>| <span data-ttu-id="12c6b-581">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-581">Attributes</span></span>| <span data-ttu-id="12c6b-582">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="12c6b-583">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-583">String</span></span>||<span data-ttu-id="12c6b-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="12c6b-586">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-586">String</span></span>||<span data-ttu-id="12c6b-587">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="12c6b-587">The subject of the item to be attached.</span></span> <span data-ttu-id="12c6b-588">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="12c6b-589">Object</span><span class="sxs-lookup"><span data-stu-id="12c6b-589">Object</span></span>| <span data-ttu-id="12c6b-590">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-590">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-591">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="12c6b-592">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-592">Object</span></span>| <span data-ttu-id="12c6b-593">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-593">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-594">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="12c6b-595">fonction</span><span class="sxs-lookup"><span data-stu-id="12c6b-595">function</span></span>| <span data-ttu-id="12c6b-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-596">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-597">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="12c6b-598">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="12c6b-599">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="12c6b-600">Erreurs</span><span class="sxs-lookup"><span data-stu-id="12c6b-600">Errors</span></span>

| <span data-ttu-id="12c6b-601">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="12c6b-601">Error code</span></span> | <span data-ttu-id="12c6b-602">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="12c6b-603">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="12c6b-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-604">Requirements</span></span>

|<span data-ttu-id="12c6b-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-605">Requirement</span></span>| <span data-ttu-id="12c6b-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-607">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-608">1.1</span><span class="sxs-lookup"><span data-stu-id="12c6b-608">1.1</span></span>|
|[<span data-ttu-id="12c6b-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="12c6b-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-612">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-613">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-613">Example</span></span>

<span data-ttu-id="12c6b-614">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="12c6b-615">close()</span><span class="sxs-lookup"><span data-stu-id="12c6b-615">close()</span></span>

<span data-ttu-id="12c6b-616">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="12c6b-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="12c6b-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-619">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="12c6b-620">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="12c6b-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-621">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-621">Requirements</span></span>

|<span data-ttu-id="12c6b-622">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-622">Requirement</span></span>| <span data-ttu-id="12c6b-623">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-624">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-625">1.3</span><span class="sxs-lookup"><span data-stu-id="12c6b-625">1.3</span></span>|
|[<span data-ttu-id="12c6b-626">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-627">Restreinte</span><span class="sxs-lookup"><span data-stu-id="12c6b-627">Restricted</span></span>|
|[<span data-ttu-id="12c6b-628">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-629">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="12c6b-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="12c6b-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="12c6b-631">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="12c6b-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-632">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="12c6b-633">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="12c6b-634">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="12c6b-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="12c6b-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-638">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-638">Parameters</span></span>

|<span data-ttu-id="12c6b-639">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-639">Name</span></span>| <span data-ttu-id="12c6b-640">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-640">Type</span></span>| <span data-ttu-id="12c6b-641">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="12c6b-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="12c6b-642">String &#124; Object</span></span>| |<span data-ttu-id="12c6b-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="12c6b-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="12c6b-645">**OR**</span></span><br/><span data-ttu-id="12c6b-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="12c6b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="12c6b-648">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-648">String</span></span> | <span data-ttu-id="12c6b-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-649">&lt;optional&gt;</span></span> | <span data-ttu-id="12c6b-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="12c6b-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="12c6b-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-653">&lt;optional&gt;</span></span> | <span data-ttu-id="12c6b-654">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="12c6b-655">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-655">String</span></span> | | <span data-ttu-id="12c6b-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="12c6b-658">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-658">String</span></span> | | <span data-ttu-id="12c6b-659">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="12c6b-660">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12c6b-660">String</span></span> | | <span data-ttu-id="12c6b-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="12c6b-663">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-663">String</span></span> | | <span data-ttu-id="12c6b-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="12c6b-667">function</span><span class="sxs-lookup"><span data-stu-id="12c6b-667">function</span></span> | <span data-ttu-id="12c6b-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-668">&lt;optional&gt;</span></span> | <span data-ttu-id="12c6b-669">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="12c6b-670">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-670">Requirements</span></span>

|<span data-ttu-id="12c6b-671">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-671">Requirement</span></span>| <span data-ttu-id="12c6b-672">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-673">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-674">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-674">1.0</span></span>|
|[<span data-ttu-id="12c6b-675">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-676">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-677">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-678">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="12c6b-679">Exemples</span><span class="sxs-lookup"><span data-stu-id="12c6b-679">Examples</span></span>

<span data-ttu-id="12c6b-680">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="12c6b-681">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="12c6b-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="12c6b-682">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="12c6b-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="12c6b-683">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="12c6b-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="12c6b-684">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="12c6b-685">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="12c6b-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="12c6b-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="12c6b-687">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="12c6b-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-688">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="12c6b-689">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="12c6b-690">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="12c6b-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="12c6b-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-694">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-694">Parameters</span></span>

|<span data-ttu-id="12c6b-695">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-695">Name</span></span>| <span data-ttu-id="12c6b-696">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-696">Type</span></span>| <span data-ttu-id="12c6b-697">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="12c6b-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="12c6b-698">String &#124; Object</span></span>| | <span data-ttu-id="12c6b-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="12c6b-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="12c6b-701">**OR**</span></span><br/><span data-ttu-id="12c6b-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="12c6b-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="12c6b-704">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-704">String</span></span> | <span data-ttu-id="12c6b-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-705">&lt;optional&gt;</span></span> | <span data-ttu-id="12c6b-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="12c6b-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="12c6b-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-709">&lt;optional&gt;</span></span> | <span data-ttu-id="12c6b-710">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="12c6b-711">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-711">String</span></span> | | <span data-ttu-id="12c6b-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="12c6b-714">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-714">String</span></span> | | <span data-ttu-id="12c6b-715">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="12c6b-716">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12c6b-716">String</span></span> | | <span data-ttu-id="12c6b-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="12c6b-719">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12c6b-719">String</span></span> | | <span data-ttu-id="12c6b-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="12c6b-723">function</span><span class="sxs-lookup"><span data-stu-id="12c6b-723">function</span></span> | <span data-ttu-id="12c6b-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-724">&lt;optional&gt;</span></span> | <span data-ttu-id="12c6b-725">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="12c6b-726">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-726">Requirements</span></span>

|<span data-ttu-id="12c6b-727">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-727">Requirement</span></span>| <span data-ttu-id="12c6b-728">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-729">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-730">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-730">1.0</span></span>|
|[<span data-ttu-id="12c6b-731">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-732">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-733">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-734">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="12c6b-735">Exemples</span><span class="sxs-lookup"><span data-stu-id="12c6b-735">Examples</span></span>

<span data-ttu-id="12c6b-736">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="12c6b-737">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="12c6b-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="12c6b-738">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="12c6b-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="12c6b-739">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="12c6b-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="12c6b-740">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="12c6b-741">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="12c6b-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="12c6b-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="12c6b-743">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="12c6b-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-744">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-745">Requirements</span></span>

|<span data-ttu-id="12c6b-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-746">Requirement</span></span>| <span data-ttu-id="12c6b-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-749">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-749">1.0</span></span>|
|[<span data-ttu-id="12c6b-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-751">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-753">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="12c6b-754">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="12c6b-754">Returns:</span></span>

<span data-ttu-id="12c6b-755">Type : [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="12c6b-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="12c6b-756">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-756">Example</span></span>

<span data-ttu-id="12c6b-757">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="12c6b-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="12c6b-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="12c6b-759">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="12c6b-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-760">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-761">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-761">Parameters</span></span>

|<span data-ttu-id="12c6b-762">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-762">Name</span></span>| <span data-ttu-id="12c6b-763">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-763">Type</span></span>| <span data-ttu-id="12c6b-764">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="12c6b-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="12c6b-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="12c6b-766">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="12c6b-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12c6b-767">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-767">Requirements</span></span>

|<span data-ttu-id="12c6b-768">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-768">Requirement</span></span>| <span data-ttu-id="12c6b-769">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-770">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-771">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-771">1.0</span></span>|
|[<span data-ttu-id="12c6b-772">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-773">Restreinte</span><span class="sxs-lookup"><span data-stu-id="12c6b-773">Restricted</span></span>|
|[<span data-ttu-id="12c6b-774">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-775">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="12c6b-776">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="12c6b-776">Returns:</span></span>

<span data-ttu-id="12c6b-777">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="12c6b-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="12c6b-778">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="12c6b-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="12c6b-779">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="12c6b-780">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="12c6b-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="12c6b-781">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="12c6b-781">Value of `entityType`</span></span> | <span data-ttu-id="12c6b-782">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="12c6b-782">Type of objects in returned array</span></span> | <span data-ttu-id="12c6b-783">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="12c6b-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="12c6b-784">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-784">String</span></span> | <span data-ttu-id="12c6b-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="12c6b-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="12c6b-786">Contact</span><span class="sxs-lookup"><span data-stu-id="12c6b-786">Contact</span></span> | <span data-ttu-id="12c6b-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="12c6b-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="12c6b-788">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-788">String</span></span> | <span data-ttu-id="12c6b-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="12c6b-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="12c6b-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="12c6b-790">MeetingSuggestion</span></span> | <span data-ttu-id="12c6b-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="12c6b-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="12c6b-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="12c6b-792">PhoneNumber</span></span> | <span data-ttu-id="12c6b-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="12c6b-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="12c6b-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="12c6b-794">TaskSuggestion</span></span> | <span data-ttu-id="12c6b-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="12c6b-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="12c6b-796">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-796">String</span></span> | <span data-ttu-id="12c6b-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="12c6b-797">**Restricted**</span></span> |

<span data-ttu-id="12c6b-798">Type : Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="12c6b-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="12c6b-799">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-799">Example</span></span>

<span data-ttu-id="12c6b-800">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="12c6b-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="12c6b-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="12c6b-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="12c6b-802">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="12c6b-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-803">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="12c6b-804">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="12c6b-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-805">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-805">Parameters</span></span>

|<span data-ttu-id="12c6b-806">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-806">Name</span></span>| <span data-ttu-id="12c6b-807">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-807">Type</span></span>| <span data-ttu-id="12c6b-808">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="12c6b-809">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-809">String</span></span>|<span data-ttu-id="12c6b-810">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="12c6b-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12c6b-811">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-811">Requirements</span></span>

|<span data-ttu-id="12c6b-812">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-812">Requirement</span></span>| <span data-ttu-id="12c6b-813">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-814">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-815">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-815">1.0</span></span>|
|[<span data-ttu-id="12c6b-816">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-817">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-818">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-819">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="12c6b-820">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="12c6b-820">Returns:</span></span>

<span data-ttu-id="12c6b-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="12c6b-823">Type : Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="12c6b-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="12c6b-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="12c6b-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="12c6b-825">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="12c6b-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-826">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="12c6b-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="12c6b-830">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="12c6b-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="12c6b-831">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="12c6b-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="12c6b-835">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-835">Requirements</span></span>

|<span data-ttu-id="12c6b-836">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-836">Requirement</span></span>| <span data-ttu-id="12c6b-837">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-838">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-839">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-839">1.0</span></span>|
|[<span data-ttu-id="12c6b-840">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-841">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-842">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-843">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="12c6b-844">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="12c6b-844">Returns:</span></span>

<span data-ttu-id="12c6b-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="12c6b-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="12c6b-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="12c6b-848">Object</span><span class="sxs-lookup"><span data-stu-id="12c6b-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="12c6b-849">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-849">Example</span></span>

<span data-ttu-id="12c6b-850">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="12c6b-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="12c6b-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="12c6b-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="12c6b-852">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="12c6b-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-853">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="12c6b-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="12c6b-854">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="12c6b-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="12c6b-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-857">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-857">Parameters</span></span>

|<span data-ttu-id="12c6b-858">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-858">Name</span></span>| <span data-ttu-id="12c6b-859">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-859">Type</span></span>| <span data-ttu-id="12c6b-860">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="12c6b-861">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-861">String</span></span>|<span data-ttu-id="12c6b-862">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="12c6b-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12c6b-863">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-863">Requirements</span></span>

|<span data-ttu-id="12c6b-864">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-864">Requirement</span></span>| <span data-ttu-id="12c6b-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-866">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-867">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-867">1.0</span></span>|
|[<span data-ttu-id="12c6b-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-869">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="12c6b-872">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="12c6b-872">Returns:</span></span>

<span data-ttu-id="12c6b-873">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="12c6b-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="12c6b-874">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="12c6b-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="12c6b-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="12c6b-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="12c6b-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="12c6b-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="12c6b-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="12c6b-878">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="12c6b-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="12c6b-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-881">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-881">Parameters</span></span>

|<span data-ttu-id="12c6b-882">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-882">Name</span></span>| <span data-ttu-id="12c6b-883">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-883">Type</span></span>| <span data-ttu-id="12c6b-884">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-884">Attributes</span></span>| <span data-ttu-id="12c6b-885">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="12c6b-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="12c6b-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="12c6b-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="12c6b-890">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-890">Object</span></span>| <span data-ttu-id="12c6b-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-891">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-892">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="12c6b-893">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-893">Object</span></span>| <span data-ttu-id="12c6b-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-894">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-895">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="12c6b-896">fonction</span><span class="sxs-lookup"><span data-stu-id="12c6b-896">function</span></span>||<span data-ttu-id="12c6b-897">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="12c6b-898">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="12c6b-899">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12c6b-900">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-900">Requirements</span></span>

|<span data-ttu-id="12c6b-901">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-901">Requirement</span></span>| <span data-ttu-id="12c6b-902">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-903">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-904">1.2</span><span class="sxs-lookup"><span data-stu-id="12c6b-904">1.2</span></span>|
|[<span data-ttu-id="12c6b-905">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="12c6b-907">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-908">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="12c6b-909">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="12c6b-909">Returns:</span></span>

<span data-ttu-id="12c6b-910">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="12c6b-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="12c6b-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="12c6b-912">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="12c6b-913">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="12c6b-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="12c6b-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="12c6b-915">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="12c6b-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="12c6b-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-919">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-919">Parameters</span></span>

|<span data-ttu-id="12c6b-920">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-920">Name</span></span>| <span data-ttu-id="12c6b-921">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-921">Type</span></span>| <span data-ttu-id="12c6b-922">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-922">Attributes</span></span>| <span data-ttu-id="12c6b-923">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="12c6b-924">function</span><span class="sxs-lookup"><span data-stu-id="12c6b-924">function</span></span>||<span data-ttu-id="12c6b-925">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="12c6b-926">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="12c6b-927">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="12c6b-928">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-928">Object</span></span>| <span data-ttu-id="12c6b-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-929">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-930">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="12c6b-931">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12c6b-932">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-932">Requirements</span></span>

|<span data-ttu-id="12c6b-933">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-933">Requirement</span></span>| <span data-ttu-id="12c6b-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-935">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-936">1.0</span><span class="sxs-lookup"><span data-stu-id="12c6b-936">1.0</span></span>|
|[<span data-ttu-id="12c6b-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-938">ReadItem</span></span>|
|[<span data-ttu-id="12c6b-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-940">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12c6b-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-941">Example</span></span>

<span data-ttu-id="12c6b-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="12c6b-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="12c6b-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="12c6b-946">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="12c6b-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="12c6b-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-951">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-951">Parameters</span></span>

|<span data-ttu-id="12c6b-952">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-952">Name</span></span>| <span data-ttu-id="12c6b-953">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-953">Type</span></span>| <span data-ttu-id="12c6b-954">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-954">Attributes</span></span>| <span data-ttu-id="12c6b-955">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="12c6b-956">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-956">String</span></span>||<span data-ttu-id="12c6b-957">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="12c6b-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="12c6b-958">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-958">Object</span></span>| <span data-ttu-id="12c6b-959">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-959">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-960">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="12c6b-961">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-961">Object</span></span>| <span data-ttu-id="12c6b-962">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-962">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-963">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="12c6b-964">fonction</span><span class="sxs-lookup"><span data-stu-id="12c6b-964">function</span></span>| <span data-ttu-id="12c6b-965">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-965">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-966">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="12c6b-967">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="12c6b-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="12c6b-968">Erreurs</span><span class="sxs-lookup"><span data-stu-id="12c6b-968">Errors</span></span>

| <span data-ttu-id="12c6b-969">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="12c6b-969">Error code</span></span> | <span data-ttu-id="12c6b-970">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="12c6b-971">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="12c6b-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="12c6b-972">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-972">Requirements</span></span>

|<span data-ttu-id="12c6b-973">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-973">Requirement</span></span>| <span data-ttu-id="12c6b-974">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-975">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-976">1.1</span><span class="sxs-lookup"><span data-stu-id="12c6b-976">1.1</span></span>|
|[<span data-ttu-id="12c6b-977">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="12c6b-979">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-980">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-981">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-981">Example</span></span>

<span data-ttu-id="12c6b-982">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="12c6b-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="12c6b-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="12c6b-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="12c6b-984">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="12c6b-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="12c6b-p166">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-988">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="12c6b-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="12c6b-989">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="12c6b-p168">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="12c6b-993">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="12c6b-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="12c6b-994">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="12c6b-994">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="12c6b-995">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="12c6b-995">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="12c6b-996">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="12c6b-996">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-997">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-997">Parameters</span></span>

|<span data-ttu-id="12c6b-998">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-998">Name</span></span>| <span data-ttu-id="12c6b-999">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-999">Type</span></span>| <span data-ttu-id="12c6b-1000">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-1000">Attributes</span></span>| <span data-ttu-id="12c6b-1001">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-1001">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="12c6b-1002">Object</span><span class="sxs-lookup"><span data-stu-id="12c6b-1002">Object</span></span>| <span data-ttu-id="12c6b-1003">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-1003">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-1004">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1004">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="12c6b-1005">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-1005">Object</span></span>| <span data-ttu-id="12c6b-1006">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-1007">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1007">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="12c6b-1008">fonction</span><span class="sxs-lookup"><span data-stu-id="12c6b-1008">function</span></span>||<span data-ttu-id="12c6b-1009">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="12c6b-1010">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1010">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12c6b-1011">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-1011">Requirements</span></span>

|<span data-ttu-id="12c6b-1012">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-1012">Requirement</span></span>| <span data-ttu-id="12c6b-1013">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-1014">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-1015">1.3</span><span class="sxs-lookup"><span data-stu-id="12c6b-1015">1.3</span></span>|
|[<span data-ttu-id="12c6b-1016">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="12c6b-1018">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-1019">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-1019">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="12c6b-1020">範例</span><span class="sxs-lookup"><span data-stu-id="12c6b-1020">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="12c6b-p170">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="12c6b-1023">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="12c6b-1023">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="12c6b-1024">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1024">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="12c6b-p171">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="12c6b-1028">Paramètres</span><span class="sxs-lookup"><span data-stu-id="12c6b-1028">Parameters</span></span>

|<span data-ttu-id="12c6b-1029">Nom</span><span class="sxs-lookup"><span data-stu-id="12c6b-1029">Name</span></span>| <span data-ttu-id="12c6b-1030">Type</span><span class="sxs-lookup"><span data-stu-id="12c6b-1030">Type</span></span>| <span data-ttu-id="12c6b-1031">Attributs</span><span class="sxs-lookup"><span data-stu-id="12c6b-1031">Attributes</span></span>| <span data-ttu-id="12c6b-1032">Description</span><span class="sxs-lookup"><span data-stu-id="12c6b-1032">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="12c6b-1033">String</span><span class="sxs-lookup"><span data-stu-id="12c6b-1033">String</span></span>||<span data-ttu-id="12c6b-p172">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="12c6b-1037">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-1037">Object</span></span>| <span data-ttu-id="12c6b-1038">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-1039">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="12c6b-1040">Objet</span><span class="sxs-lookup"><span data-stu-id="12c6b-1040">Object</span></span>| <span data-ttu-id="12c6b-1041">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-1042">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="12c6b-1043">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="12c6b-1043">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="12c6b-1044">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="12c6b-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="12c6b-p173">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="12c6b-p174">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="12c6b-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="12c6b-1049">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="12c6b-1049">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="12c6b-1050">fonction</span><span class="sxs-lookup"><span data-stu-id="12c6b-1050">function</span></span>||<span data-ttu-id="12c6b-1051">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="12c6b-1051">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="12c6b-1052">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12c6b-1052">Requirements</span></span>

|<span data-ttu-id="12c6b-1053">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12c6b-1053">Requirement</span></span>| <span data-ttu-id="12c6b-1054">Valeur</span><span class="sxs-lookup"><span data-stu-id="12c6b-1054">Value</span></span>|
|---|---|
|[<span data-ttu-id="12c6b-1055">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12c6b-1055">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12c6b-1056">1.2</span><span class="sxs-lookup"><span data-stu-id="12c6b-1056">1.2</span></span>|
|[<span data-ttu-id="12c6b-1057">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="12c6b-1057">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12c6b-1058">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="12c6b-1058">ReadWriteItem</span></span>|
|[<span data-ttu-id="12c6b-1059">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12c6b-1059">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12c6b-1060">Composition</span><span class="sxs-lookup"><span data-stu-id="12c6b-1060">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="12c6b-1061">Exemple</span><span class="sxs-lookup"><span data-stu-id="12c6b-1061">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
