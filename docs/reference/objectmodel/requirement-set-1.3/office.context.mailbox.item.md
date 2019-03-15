---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 6896c849b144f3720a3d9fb284e88d18d5c8ff43
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600297"
---
# <a name="item"></a><span data-ttu-id="d20b8-102">élément</span><span class="sxs-lookup"><span data-stu-id="d20b8-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d20b8-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d20b8-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d20b8-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="d20b8-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-106">Requirements</span></span>

|<span data-ttu-id="d20b8-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-107">Requirement</span></span>| <span data-ttu-id="d20b8-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-110">1.0</span></span>|
|[<span data-ttu-id="d20b8-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d20b8-112">Restricted</span></span>|
|[<span data-ttu-id="d20b8-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="d20b8-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-115">Example</span></span>

<span data-ttu-id="d20b8-116">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="d20b8-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d20b8-117">Membres</span><span class="sxs-lookup"><span data-stu-id="d20b8-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="d20b8-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d20b8-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="d20b8-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-121">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="d20b8-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d20b8-122">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="d20b8-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-123">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-123">Type</span></span>

*   <span data-ttu-id="d20b8-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d20b8-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-125">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-125">Requirements</span></span>

|<span data-ttu-id="d20b8-126">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-126">Requirement</span></span>| <span data-ttu-id="d20b8-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-128">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-129">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-129">1.0</span></span>|
|[<span data-ttu-id="d20b8-130">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-131">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-134">Example</span></span>

<span data-ttu-id="d20b8-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d20b8-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d20b8-137">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="d20b8-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d20b8-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-139">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-139">Type</span></span>

*   [<span data-ttu-id="d20b8-140">Destinataires</span><span class="sxs-lookup"><span data-stu-id="d20b8-140">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d20b8-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-141">Requirements</span></span>

|<span data-ttu-id="d20b8-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-142">Requirement</span></span>| <span data-ttu-id="d20b8-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-145">1.1</span><span class="sxs-lookup"><span data-stu-id="d20b8-145">1.1</span></span>|
|[<span data-ttu-id="d20b8-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-147">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-149">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="d20b8-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="d20b8-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="d20b8-152">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-153">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-153">Type</span></span>

*   [<span data-ttu-id="d20b8-154">Body</span><span class="sxs-lookup"><span data-stu-id="d20b8-154">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="d20b8-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-155">Requirements</span></span>

|<span data-ttu-id="d20b8-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-156">Requirement</span></span>| <span data-ttu-id="d20b8-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-159">1.1</span><span class="sxs-lookup"><span data-stu-id="d20b8-159">1.1</span></span>|
|[<span data-ttu-id="d20b8-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-161">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-164">Example</span></span>

<span data-ttu-id="d20b8-165">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="d20b8-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d20b8-166">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d20b8-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d20b8-168">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="d20b8-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d20b8-169">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-170">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-170">Read mode</span></span>

<span data-ttu-id="d20b8-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-173">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-173">Compose mode</span></span>

<span data-ttu-id="d20b8-174">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="d20b8-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d20b8-175">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-175">Type</span></span>

*   <span data-ttu-id="d20b8-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-177">Requirements</span></span>

|<span data-ttu-id="d20b8-178">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-178">Requirement</span></span>| <span data-ttu-id="d20b8-179">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-180">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-181">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-181">1.0</span></span>|
|[<span data-ttu-id="d20b8-182">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-182">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-183">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-184">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-185">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="d20b8-186">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="d20b8-186">(nullable) conversationId :String</span></span>

<span data-ttu-id="d20b8-187">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="d20b8-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d20b8-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d20b8-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-192">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-192">Type</span></span>

*   <span data-ttu-id="d20b8-193">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-194">Requirements</span></span>

|<span data-ttu-id="d20b8-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-195">Requirement</span></span>| <span data-ttu-id="d20b8-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-198">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-198">1.0</span></span>|
|[<span data-ttu-id="d20b8-199">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-199">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-200">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-201">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="d20b8-204">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="d20b8-204">dateTimeCreated :Date</span></span>

<span data-ttu-id="d20b8-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-207">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-207">Type</span></span>

*   <span data-ttu-id="d20b8-208">Date</span><span class="sxs-lookup"><span data-stu-id="d20b8-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-209">Requirements</span></span>

|<span data-ttu-id="d20b8-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-210">Requirement</span></span>| <span data-ttu-id="d20b8-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-213">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-213">1.0</span></span>|
|[<span data-ttu-id="d20b8-214">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-215">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-217">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-218">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="d20b8-219">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="d20b8-219">dateTimeModified :Date</span></span>

<span data-ttu-id="d20b8-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-222">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-223">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-223">Type</span></span>

*   <span data-ttu-id="d20b8-224">Date</span><span class="sxs-lookup"><span data-stu-id="d20b8-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-225">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-225">Requirements</span></span>

|<span data-ttu-id="d20b8-226">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-226">Requirement</span></span>| <span data-ttu-id="d20b8-227">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-228">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-229">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-229">1.0</span></span>|
|[<span data-ttu-id="d20b8-230">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-230">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-231">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-232">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-232">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-233">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-234">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="d20b8-235">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d20b8-235">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="d20b8-236">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d20b8-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-239">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-239">Read mode</span></span>

<span data-ttu-id="d20b8-240">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-241">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-241">Compose mode</span></span>

<span data-ttu-id="d20b8-242">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d20b8-243">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d20b8-244">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d20b8-245">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-245">Type</span></span>

*   <span data-ttu-id="d20b8-246">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d20b8-246">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-247">Requirements</span></span>

|<span data-ttu-id="d20b8-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-248">Requirement</span></span>| <span data-ttu-id="d20b8-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-251">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-251">1.0</span></span>|
|[<span data-ttu-id="d20b8-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-252">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-253">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-254">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="d20b8-256">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d20b8-256">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="d20b8-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="d20b8-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-261">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-262">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-262">Type</span></span>

*   [<span data-ttu-id="d20b8-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d20b8-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d20b8-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-264">Requirements</span></span>

|<span data-ttu-id="d20b8-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-265">Requirement</span></span>| <span data-ttu-id="d20b8-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-268">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-268">1.0</span></span>|
|[<span data-ttu-id="d20b8-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-270">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-272">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-273">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="d20b8-274">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="d20b8-274">internetMessageId :String</span></span>

<span data-ttu-id="d20b8-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-277">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-277">Type</span></span>

*   <span data-ttu-id="d20b8-278">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-279">Requirements</span></span>

|<span data-ttu-id="d20b8-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-280">Requirement</span></span>| <span data-ttu-id="d20b8-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-283">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-283">1.0</span></span>|
|[<span data-ttu-id="d20b8-284">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-285">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-287">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="d20b8-289">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="d20b8-289">itemClass :String</span></span>

<span data-ttu-id="d20b8-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d20b8-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="d20b8-294">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-294">Type</span></span> | <span data-ttu-id="d20b8-295">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-295">Description</span></span> | <span data-ttu-id="d20b8-296">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="d20b8-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="d20b8-297">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="d20b8-297">Appointment items</span></span> | <span data-ttu-id="d20b8-298">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="d20b8-299">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="d20b8-299">Message items</span></span> | <span data-ttu-id="d20b8-300">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="d20b8-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="d20b8-301">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-302">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-302">Type</span></span>

*   <span data-ttu-id="d20b8-303">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-304">Requirements</span></span>

|<span data-ttu-id="d20b8-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-305">Requirement</span></span>| <span data-ttu-id="d20b8-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-308">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-308">1.0</span></span>|
|[<span data-ttu-id="d20b8-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-310">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d20b8-314">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="d20b8-314">(nullable) itemId :String</span></span>

<span data-ttu-id="d20b8-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-317">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="d20b8-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d20b8-318">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="d20b8-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d20b8-319">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="d20b8-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d20b8-320">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="d20b8-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d20b8-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-323">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-323">Type</span></span>

*   <span data-ttu-id="d20b8-324">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-325">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-325">Requirements</span></span>

|<span data-ttu-id="d20b8-326">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-326">Requirement</span></span>| <span data-ttu-id="d20b8-327">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-328">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-329">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-329">1.0</span></span>|
|[<span data-ttu-id="d20b8-330">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-331">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-332">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-333">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-334">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-334">Example</span></span>

<span data-ttu-id="d20b8-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="d20b8-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d20b8-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d20b8-338">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="d20b8-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d20b8-339">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-340">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-340">Type</span></span>

*   [<span data-ttu-id="d20b8-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d20b8-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d20b8-342">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-342">Requirements</span></span>

|<span data-ttu-id="d20b8-343">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-343">Requirement</span></span>| <span data-ttu-id="d20b8-344">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-345">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-346">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-346">1.0</span></span>|
|[<span data-ttu-id="d20b8-347">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-347">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-348">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-349">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-349">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-350">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-351">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="d20b8-352">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="d20b8-352">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="d20b8-353">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-354">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-354">Read mode</span></span>

<span data-ttu-id="d20b8-355">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-356">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-356">Compose mode</span></span>

<span data-ttu-id="d20b8-357">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d20b8-358">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-358">Type</span></span>

*   <span data-ttu-id="d20b8-359">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="d20b8-359">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-360">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-360">Requirements</span></span>

|<span data-ttu-id="d20b8-361">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-361">Requirement</span></span>| <span data-ttu-id="d20b8-362">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-363">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-364">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-364">1.0</span></span>|
|[<span data-ttu-id="d20b8-365">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-366">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-367">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-368">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d20b8-369">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="d20b8-369">normalizedSubject :String</span></span>

<span data-ttu-id="d20b8-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d20b8-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="d20b8-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-374">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-374">Type</span></span>

*   <span data-ttu-id="d20b8-375">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-376">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-376">Requirements</span></span>

|<span data-ttu-id="d20b8-377">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-377">Requirement</span></span>| <span data-ttu-id="d20b8-378">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-379">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-380">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-380">1.0</span></span>|
|[<span data-ttu-id="d20b8-381">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-381">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-382">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-383">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-383">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-384">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-385">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="d20b8-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d20b8-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="d20b8-387">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-388">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-388">Type</span></span>

*   [<span data-ttu-id="d20b8-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d20b8-389">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d20b8-390">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-390">Requirements</span></span>

|<span data-ttu-id="d20b8-391">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-391">Requirement</span></span>| <span data-ttu-id="d20b8-392">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-393">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-394">1.3</span><span class="sxs-lookup"><span data-stu-id="d20b8-394">1.3</span></span>|
|[<span data-ttu-id="d20b8-395">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-395">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-396">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-397">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-398">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-399">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d20b8-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d20b8-401">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d20b8-402">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-403">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-403">Read mode</span></span>

<span data-ttu-id="d20b8-404">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="d20b8-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-405">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-405">Compose mode</span></span>

<span data-ttu-id="d20b8-406">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="d20b8-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d20b8-407">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-407">Type</span></span>

*   <span data-ttu-id="d20b8-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-409">Requirements</span></span>

|<span data-ttu-id="d20b8-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-410">Requirement</span></span>| <span data-ttu-id="d20b8-411">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-412">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-413">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-413">1.0</span></span>|
|[<span data-ttu-id="d20b8-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-414">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-415">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-416">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-416">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-417">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="d20b8-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d20b8-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="d20b8-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-421">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-421">Type</span></span>

*   [<span data-ttu-id="d20b8-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d20b8-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d20b8-423">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-423">Requirements</span></span>

|<span data-ttu-id="d20b8-424">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-424">Requirement</span></span>| <span data-ttu-id="d20b8-425">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-426">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-427">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-427">1.0</span></span>|
|[<span data-ttu-id="d20b8-428">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-429">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-431">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-432">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d20b8-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d20b8-434">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d20b8-435">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-436">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-436">Read mode</span></span>

<span data-ttu-id="d20b8-437">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="d20b8-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-438">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-438">Compose mode</span></span>

<span data-ttu-id="d20b8-439">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="d20b8-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d20b8-440">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-440">Type</span></span>

*   <span data-ttu-id="d20b8-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-442">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-442">Requirements</span></span>

|<span data-ttu-id="d20b8-443">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-443">Requirement</span></span>| <span data-ttu-id="d20b8-444">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-445">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-446">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-446">1.0</span></span>|
|[<span data-ttu-id="d20b8-447">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-448">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-449">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-450">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="d20b8-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d20b8-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="d20b8-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d20b8-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-456">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d20b8-457">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-457">Type</span></span>

*   [<span data-ttu-id="d20b8-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d20b8-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d20b8-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-459">Requirements</span></span>

|<span data-ttu-id="d20b8-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-460">Requirement</span></span>| <span data-ttu-id="d20b8-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-463">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-463">1.0</span></span>|
|[<span data-ttu-id="d20b8-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-465">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-467">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="d20b8-469">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d20b8-469">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="d20b8-470">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d20b8-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-473">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-473">Read mode</span></span>

<span data-ttu-id="d20b8-474">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-475">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-475">Compose mode</span></span>

<span data-ttu-id="d20b8-476">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d20b8-477">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d20b8-478">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d20b8-479">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-479">Type</span></span>

*   <span data-ttu-id="d20b8-480">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="d20b8-480">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-481">Requirements</span></span>

|<span data-ttu-id="d20b8-482">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-482">Requirement</span></span>| <span data-ttu-id="d20b8-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-485">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-485">1.0</span></span>|
|[<span data-ttu-id="d20b8-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-487">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-489">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-489">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="d20b8-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d20b8-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="d20b8-491">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d20b8-492">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="d20b8-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-493">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-493">Read mode</span></span>

<span data-ttu-id="d20b8-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-496">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-496">Compose mode</span></span>

<span data-ttu-id="d20b8-497">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="d20b8-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d20b8-498">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-498">Type</span></span>

*   <span data-ttu-id="d20b8-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d20b8-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-500">Requirements</span></span>

|<span data-ttu-id="d20b8-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-501">Requirement</span></span>| <span data-ttu-id="d20b8-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-504">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-504">1.0</span></span>|
|[<span data-ttu-id="d20b8-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-506">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-508">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-508">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="d20b8-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="d20b8-510">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="d20b8-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d20b8-511">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d20b8-512">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-512">Read mode</span></span>

<span data-ttu-id="d20b8-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d20b8-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-515">Compose mode</span></span>

<span data-ttu-id="d20b8-516">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="d20b8-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d20b8-517">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-517">Type</span></span>

*   <span data-ttu-id="d20b8-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d20b8-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-519">Requirements</span></span>

|<span data-ttu-id="d20b8-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-520">Requirement</span></span>| <span data-ttu-id="d20b8-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-523">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-523">1.0</span></span>|
|[<span data-ttu-id="d20b8-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-525">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-527">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d20b8-528">Méthodes</span><span class="sxs-lookup"><span data-stu-id="d20b8-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d20b8-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d20b8-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d20b8-530">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="d20b8-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d20b8-531">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="d20b8-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d20b8-532">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="d20b8-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-533">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-533">Parameters</span></span>

|<span data-ttu-id="d20b8-534">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-534">Name</span></span>| <span data-ttu-id="d20b8-535">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-535">Type</span></span>| <span data-ttu-id="d20b8-536">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-536">Attributes</span></span>| <span data-ttu-id="d20b8-537">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="d20b8-538">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-538">String</span></span>||<span data-ttu-id="d20b8-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d20b8-541">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-541">String</span></span>||<span data-ttu-id="d20b8-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d20b8-544">Object</span><span class="sxs-lookup"><span data-stu-id="d20b8-544">Object</span></span>| <span data-ttu-id="d20b8-545">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-545">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-546">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d20b8-547">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-547">Object</span></span>| <span data-ttu-id="d20b8-548">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-548">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-549">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d20b8-550">fonction</span><span class="sxs-lookup"><span data-stu-id="d20b8-550">function</span></span>| <span data-ttu-id="d20b8-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-551">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-552">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d20b8-553">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d20b8-554">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d20b8-555">Erreurs</span><span class="sxs-lookup"><span data-stu-id="d20b8-555">Errors</span></span>

| <span data-ttu-id="d20b8-556">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="d20b8-556">Error code</span></span> | <span data-ttu-id="d20b8-557">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="d20b8-558">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="d20b8-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="d20b8-559">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="d20b8-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d20b8-560">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d20b8-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-561">Requirements</span></span>

|<span data-ttu-id="d20b8-562">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-562">Requirement</span></span>| <span data-ttu-id="d20b8-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-565">1.1</span><span class="sxs-lookup"><span data-stu-id="d20b8-565">1.1</span></span>|
|[<span data-ttu-id="d20b8-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="d20b8-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-569">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d20b8-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d20b8-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d20b8-572">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d20b8-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d20b8-576">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="d20b8-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d20b8-577">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="d20b8-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-578">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-578">Parameters</span></span>

|<span data-ttu-id="d20b8-579">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-579">Name</span></span>| <span data-ttu-id="d20b8-580">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-580">Type</span></span>| <span data-ttu-id="d20b8-581">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-581">Attributes</span></span>| <span data-ttu-id="d20b8-582">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="d20b8-583">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-583">String</span></span>||<span data-ttu-id="d20b8-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d20b8-586">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-586">String</span></span>||<span data-ttu-id="d20b8-587">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="d20b8-587">The subject of the item to be attached.</span></span> <span data-ttu-id="d20b8-588">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d20b8-589">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-589">Object</span></span>| <span data-ttu-id="d20b8-590">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-590">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-591">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d20b8-592">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-592">Object</span></span>| <span data-ttu-id="d20b8-593">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-593">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-594">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d20b8-595">fonction</span><span class="sxs-lookup"><span data-stu-id="d20b8-595">function</span></span>| <span data-ttu-id="d20b8-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-596">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-597">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d20b8-598">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d20b8-599">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d20b8-600">Erreurs</span><span class="sxs-lookup"><span data-stu-id="d20b8-600">Errors</span></span>

| <span data-ttu-id="d20b8-601">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="d20b8-601">Error code</span></span> | <span data-ttu-id="d20b8-602">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d20b8-603">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d20b8-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-604">Requirements</span></span>

|<span data-ttu-id="d20b8-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-605">Requirement</span></span>| <span data-ttu-id="d20b8-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-607">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-608">1.1</span><span class="sxs-lookup"><span data-stu-id="d20b8-608">1.1</span></span>|
|[<span data-ttu-id="d20b8-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="d20b8-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-612">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-613">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-613">Example</span></span>

<span data-ttu-id="d20b8-614">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="d20b8-615">close()</span><span class="sxs-lookup"><span data-stu-id="d20b8-615">close()</span></span>

<span data-ttu-id="d20b8-616">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="d20b8-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d20b8-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-619">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d20b8-620">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="d20b8-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-621">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-621">Requirements</span></span>

|<span data-ttu-id="d20b8-622">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-622">Requirement</span></span>| <span data-ttu-id="d20b8-623">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-624">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-625">1.3</span><span class="sxs-lookup"><span data-stu-id="d20b8-625">1.3</span></span>|
|[<span data-ttu-id="d20b8-626">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-627">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d20b8-627">Restricted</span></span>|
|[<span data-ttu-id="d20b8-628">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-629">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d20b8-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d20b8-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d20b8-631">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d20b8-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-632">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d20b8-633">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d20b8-634">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="d20b8-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d20b8-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-638">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-638">Parameters</span></span>

|<span data-ttu-id="d20b8-639">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-639">Name</span></span>| <span data-ttu-id="d20b8-640">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-640">Type</span></span>| <span data-ttu-id="d20b8-641">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d20b8-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d20b8-642">String &#124; Object</span></span>| |<span data-ttu-id="d20b8-643">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="d20b8-643">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="d20b8-644">La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d20b8-644">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d20b8-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="d20b8-645">**OR**</span></span><br/><span data-ttu-id="d20b8-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="d20b8-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d20b8-648">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-648">String</span></span> | <span data-ttu-id="d20b8-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-649">&lt;optional&gt;</span></span> | <span data-ttu-id="d20b8-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d20b8-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d20b8-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-653">&lt;optional&gt;</span></span> | <span data-ttu-id="d20b8-654">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d20b8-655">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-655">String</span></span> | | <span data-ttu-id="d20b8-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d20b8-658">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-658">String</span></span> | | <span data-ttu-id="d20b8-659">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d20b8-660">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d20b8-660">String</span></span> | | <span data-ttu-id="d20b8-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d20b8-663">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d20b8-663">String</span></span> | | <span data-ttu-id="d20b8-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d20b8-667">function</span><span class="sxs-lookup"><span data-stu-id="d20b8-667">function</span></span> | <span data-ttu-id="d20b8-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-668">&lt;optional&gt;</span></span> | <span data-ttu-id="d20b8-669">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d20b8-670">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-670">Requirements</span></span>

|<span data-ttu-id="d20b8-671">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-671">Requirement</span></span>| <span data-ttu-id="d20b8-672">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-673">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-674">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-674">1.0</span></span>|
|[<span data-ttu-id="d20b8-675">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-676">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-677">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-678">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d20b8-679">Exemples</span><span class="sxs-lookup"><span data-stu-id="d20b8-679">Examples</span></span>

<span data-ttu-id="d20b8-680">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d20b8-681">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="d20b8-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d20b8-682">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="d20b8-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d20b8-683">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="d20b8-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d20b8-684">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d20b8-685">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d20b8-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d20b8-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d20b8-687">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d20b8-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-688">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d20b8-689">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d20b8-690">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="d20b8-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d20b8-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-694">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-694">Parameters</span></span>

|<span data-ttu-id="d20b8-695">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-695">Name</span></span>| <span data-ttu-id="d20b8-696">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-696">Type</span></span>| <span data-ttu-id="d20b8-697">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d20b8-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d20b8-698">String &#124; Object</span></span>| | <span data-ttu-id="d20b8-699">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse.</span><span class="sxs-lookup"><span data-stu-id="d20b8-699">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="d20b8-700">La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d20b8-700">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d20b8-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="d20b8-701">**OR**</span></span><br/><span data-ttu-id="d20b8-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="d20b8-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d20b8-704">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-704">String</span></span> | <span data-ttu-id="d20b8-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-705">&lt;optional&gt;</span></span> | <span data-ttu-id="d20b8-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d20b8-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d20b8-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-709">&lt;optional&gt;</span></span> | <span data-ttu-id="d20b8-710">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d20b8-711">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-711">String</span></span> | | <span data-ttu-id="d20b8-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d20b8-714">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-714">String</span></span> | | <span data-ttu-id="d20b8-715">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d20b8-716">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-716">String</span></span> | | <span data-ttu-id="d20b8-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d20b8-719">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d20b8-719">String</span></span> | | <span data-ttu-id="d20b8-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d20b8-723">function</span><span class="sxs-lookup"><span data-stu-id="d20b8-723">function</span></span> | <span data-ttu-id="d20b8-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-724">&lt;optional&gt;</span></span> | <span data-ttu-id="d20b8-725">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d20b8-726">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-726">Requirements</span></span>

|<span data-ttu-id="d20b8-727">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-727">Requirement</span></span>| <span data-ttu-id="d20b8-728">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-729">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-730">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-730">1.0</span></span>|
|[<span data-ttu-id="d20b8-731">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-732">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-733">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-734">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d20b8-735">Exemples</span><span class="sxs-lookup"><span data-stu-id="d20b8-735">Examples</span></span>

<span data-ttu-id="d20b8-736">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d20b8-737">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="d20b8-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d20b8-738">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="d20b8-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d20b8-739">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="d20b8-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d20b8-740">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d20b8-741">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="d20b8-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d20b8-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="d20b8-743">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d20b8-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-744">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-745">Requirements</span></span>

|<span data-ttu-id="d20b8-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-746">Requirement</span></span>| <span data-ttu-id="d20b8-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-749">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-749">1.0</span></span>|
|[<span data-ttu-id="d20b8-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-751">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-753">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d20b8-754">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d20b8-754">Returns:</span></span>

<span data-ttu-id="d20b8-755">Type : [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d20b8-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d20b8-756">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-756">Example</span></span>

<span data-ttu-id="d20b8-757">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="d20b8-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d20b8-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d20b8-759">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d20b8-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-760">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-761">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-761">Parameters</span></span>

|<span data-ttu-id="d20b8-762">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-762">Name</span></span>| <span data-ttu-id="d20b8-763">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-763">Type</span></span>| <span data-ttu-id="d20b8-764">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="d20b8-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d20b8-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="d20b8-766">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="d20b8-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d20b8-767">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-767">Requirements</span></span>

|<span data-ttu-id="d20b8-768">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-768">Requirement</span></span>| <span data-ttu-id="d20b8-769">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-770">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-771">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-771">1.0</span></span>|
|[<span data-ttu-id="d20b8-772">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-773">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d20b8-773">Restricted</span></span>|
|[<span data-ttu-id="d20b8-774">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-775">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d20b8-776">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d20b8-776">Returns:</span></span>

<span data-ttu-id="d20b8-777">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="d20b8-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d20b8-778">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="d20b8-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d20b8-779">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d20b8-780">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="d20b8-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="d20b8-781">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="d20b8-781">Value of `entityType`</span></span> | <span data-ttu-id="d20b8-782">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="d20b8-782">Type of objects in returned array</span></span> | <span data-ttu-id="d20b8-783">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="d20b8-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="d20b8-784">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-784">String</span></span> | <span data-ttu-id="d20b8-785">**Restreinte**</span><span class="sxs-lookup"><span data-stu-id="d20b8-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="d20b8-786">Contact</span><span class="sxs-lookup"><span data-stu-id="d20b8-786">Contact</span></span> | <span data-ttu-id="d20b8-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d20b8-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="d20b8-788">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-788">String</span></span> | <span data-ttu-id="d20b8-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d20b8-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="d20b8-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d20b8-790">MeetingSuggestion</span></span> | <span data-ttu-id="d20b8-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d20b8-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="d20b8-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d20b8-792">PhoneNumber</span></span> | <span data-ttu-id="d20b8-793">**Restreinte**</span><span class="sxs-lookup"><span data-stu-id="d20b8-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="d20b8-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d20b8-794">TaskSuggestion</span></span> | <span data-ttu-id="d20b8-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d20b8-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="d20b8-796">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-796">String</span></span> | <span data-ttu-id="d20b8-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d20b8-797">**Restricted**</span></span> |

<span data-ttu-id="d20b8-798">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d20b8-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d20b8-799">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-799">Example</span></span>

<span data-ttu-id="d20b8-800">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="d20b8-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="d20b8-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d20b8-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d20b8-802">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d20b8-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-803">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d20b8-804">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="d20b8-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-805">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-805">Parameters</span></span>

|<span data-ttu-id="d20b8-806">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-806">Name</span></span>| <span data-ttu-id="d20b8-807">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-807">Type</span></span>| <span data-ttu-id="d20b8-808">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d20b8-809">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-809">String</span></span>|<span data-ttu-id="d20b8-810">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="d20b8-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d20b8-811">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-811">Requirements</span></span>

|<span data-ttu-id="d20b8-812">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-812">Requirement</span></span>| <span data-ttu-id="d20b8-813">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-814">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-815">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-815">1.0</span></span>|
|[<span data-ttu-id="d20b8-816">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-817">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-818">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-819">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d20b8-820">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d20b8-820">Returns:</span></span>

<span data-ttu-id="d20b8-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d20b8-823">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d20b8-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="d20b8-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d20b8-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d20b8-825">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d20b8-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-826">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d20b8-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d20b8-830">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="d20b8-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d20b8-831">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d20b8-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d20b8-835">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-835">Requirements</span></span>

|<span data-ttu-id="d20b8-836">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-836">Requirement</span></span>| <span data-ttu-id="d20b8-837">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-838">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-839">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-839">1.0</span></span>|
|[<span data-ttu-id="d20b8-840">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-841">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-842">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-843">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d20b8-844">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d20b8-844">Returns:</span></span>

<span data-ttu-id="d20b8-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d20b8-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="d20b8-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d20b8-848">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d20b8-849">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-849">Example</span></span>

<span data-ttu-id="d20b8-850">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="d20b8-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d20b8-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d20b8-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d20b8-852">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d20b8-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-853">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="d20b8-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d20b8-854">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="d20b8-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d20b8-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-857">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-857">Parameters</span></span>

|<span data-ttu-id="d20b8-858">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-858">Name</span></span>| <span data-ttu-id="d20b8-859">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-859">Type</span></span>| <span data-ttu-id="d20b8-860">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d20b8-861">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-861">String</span></span>|<span data-ttu-id="d20b8-862">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="d20b8-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d20b8-863">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-863">Requirements</span></span>

|<span data-ttu-id="d20b8-864">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-864">Requirement</span></span>| <span data-ttu-id="d20b8-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-866">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-867">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-867">1.0</span></span>|
|[<span data-ttu-id="d20b8-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-869">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d20b8-872">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d20b8-872">Returns:</span></span>

<span data-ttu-id="d20b8-873">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d20b8-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="d20b8-874">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="d20b8-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d20b8-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d20b8-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d20b8-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d20b8-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d20b8-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d20b8-878">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d20b8-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d20b8-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-881">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-881">Parameters</span></span>

|<span data-ttu-id="d20b8-882">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-882">Name</span></span>| <span data-ttu-id="d20b8-883">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-883">Type</span></span>| <span data-ttu-id="d20b8-884">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-884">Attributes</span></span>| <span data-ttu-id="d20b8-885">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="d20b8-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d20b8-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d20b8-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="d20b8-890">Object</span><span class="sxs-lookup"><span data-stu-id="d20b8-890">Object</span></span>| <span data-ttu-id="d20b8-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-891">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-892">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d20b8-893">Object</span><span class="sxs-lookup"><span data-stu-id="d20b8-893">Object</span></span>| <span data-ttu-id="d20b8-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-894">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-895">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d20b8-896">fonction</span><span class="sxs-lookup"><span data-stu-id="d20b8-896">function</span></span>||<span data-ttu-id="d20b8-897">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d20b8-898">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d20b8-899">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d20b8-900">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-900">Requirements</span></span>

|<span data-ttu-id="d20b8-901">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-901">Requirement</span></span>| <span data-ttu-id="d20b8-902">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-903">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-904">1.2</span><span class="sxs-lookup"><span data-stu-id="d20b8-904">1.2</span></span>|
|[<span data-ttu-id="d20b8-905">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="d20b8-907">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-908">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d20b8-909">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="d20b8-909">Returns:</span></span>

<span data-ttu-id="d20b8-910">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="d20b8-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="d20b8-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d20b8-912">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d20b8-913">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d20b8-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d20b8-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d20b8-915">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d20b8-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d20b8-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-919">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-919">Parameters</span></span>

|<span data-ttu-id="d20b8-920">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-920">Name</span></span>| <span data-ttu-id="d20b8-921">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-921">Type</span></span>| <span data-ttu-id="d20b8-922">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-922">Attributes</span></span>| <span data-ttu-id="d20b8-923">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d20b8-924">function</span><span class="sxs-lookup"><span data-stu-id="d20b8-924">function</span></span>||<span data-ttu-id="d20b8-925">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d20b8-926">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d20b8-927">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="d20b8-928">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-928">Object</span></span>| <span data-ttu-id="d20b8-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-929">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-930">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d20b8-931">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d20b8-932">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-932">Requirements</span></span>

|<span data-ttu-id="d20b8-933">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-933">Requirement</span></span>| <span data-ttu-id="d20b8-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-935">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-936">1.0</span><span class="sxs-lookup"><span data-stu-id="d20b8-936">1.0</span></span>|
|[<span data-ttu-id="d20b8-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-938">ReadItem</span></span>|
|[<span data-ttu-id="d20b8-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-940">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d20b8-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-941">Example</span></span>

<span data-ttu-id="d20b8-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d20b8-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d20b8-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d20b8-946">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="d20b8-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d20b8-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-951">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-951">Parameters</span></span>

|<span data-ttu-id="d20b8-952">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-952">Name</span></span>| <span data-ttu-id="d20b8-953">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-953">Type</span></span>| <span data-ttu-id="d20b8-954">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-954">Attributes</span></span>| <span data-ttu-id="d20b8-955">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="d20b8-956">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-956">String</span></span>||<span data-ttu-id="d20b8-957">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="d20b8-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="d20b8-958">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-958">Object</span></span>| <span data-ttu-id="d20b8-959">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-959">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-960">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d20b8-961">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-961">Object</span></span>| <span data-ttu-id="d20b8-962">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-962">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-963">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d20b8-964">fonction</span><span class="sxs-lookup"><span data-stu-id="d20b8-964">function</span></span>| <span data-ttu-id="d20b8-965">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-965">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-966">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d20b8-967">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="d20b8-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d20b8-968">Erreurs</span><span class="sxs-lookup"><span data-stu-id="d20b8-968">Errors</span></span>

| <span data-ttu-id="d20b8-969">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="d20b8-969">Error code</span></span> | <span data-ttu-id="d20b8-970">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="d20b8-971">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="d20b8-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d20b8-972">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-972">Requirements</span></span>

|<span data-ttu-id="d20b8-973">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-973">Requirement</span></span>| <span data-ttu-id="d20b8-974">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-975">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-976">1.1</span><span class="sxs-lookup"><span data-stu-id="d20b8-976">1.1</span></span>|
|[<span data-ttu-id="d20b8-977">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-977">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="d20b8-979">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-979">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-980">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-981">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-981">Example</span></span>

<span data-ttu-id="d20b8-982">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="d20b8-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="d20b8-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d20b8-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="d20b8-984">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d20b8-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="d20b8-p166">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-988">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="d20b8-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d20b8-989">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d20b8-p168">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d20b8-993">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="d20b8-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d20b8-994">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="d20b8-994">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="d20b8-995">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="d20b8-995">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="d20b8-996">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="d20b8-996">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-997">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-997">Parameters</span></span>

|<span data-ttu-id="d20b8-998">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-998">Name</span></span>| <span data-ttu-id="d20b8-999">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-999">Type</span></span>| <span data-ttu-id="d20b8-1000">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-1000">Attributes</span></span>| <span data-ttu-id="d20b8-1001">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-1001">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="d20b8-1002">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-1002">Object</span></span>| <span data-ttu-id="d20b8-1003">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-1003">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-1004">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1004">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d20b8-1005">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-1005">Object</span></span>| <span data-ttu-id="d20b8-1006">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-1007">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1007">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d20b8-1008">fonction</span><span class="sxs-lookup"><span data-stu-id="d20b8-1008">function</span></span>||<span data-ttu-id="d20b8-1009">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d20b8-1010">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1010">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d20b8-1011">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-1011">Requirements</span></span>

|<span data-ttu-id="d20b8-1012">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-1012">Requirement</span></span>| <span data-ttu-id="d20b8-1013">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-1014">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-1015">1.3</span><span class="sxs-lookup"><span data-stu-id="d20b8-1015">1.3</span></span>|
|[<span data-ttu-id="d20b8-1016">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-1016">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="d20b8-1018">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-1018">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-1019">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-1019">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d20b8-1020">範例</span><span class="sxs-lookup"><span data-stu-id="d20b8-1020">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d20b8-p170">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d20b8-1023">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d20b8-1023">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d20b8-1024">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1024">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d20b8-p171">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d20b8-1028">Paramètres</span><span class="sxs-lookup"><span data-stu-id="d20b8-1028">Parameters</span></span>

|<span data-ttu-id="d20b8-1029">Nom</span><span class="sxs-lookup"><span data-stu-id="d20b8-1029">Name</span></span>| <span data-ttu-id="d20b8-1030">Type</span><span class="sxs-lookup"><span data-stu-id="d20b8-1030">Type</span></span>| <span data-ttu-id="d20b8-1031">Attributs</span><span class="sxs-lookup"><span data-stu-id="d20b8-1031">Attributes</span></span>| <span data-ttu-id="d20b8-1032">Description</span><span class="sxs-lookup"><span data-stu-id="d20b8-1032">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d20b8-1033">String</span><span class="sxs-lookup"><span data-stu-id="d20b8-1033">String</span></span>||<span data-ttu-id="d20b8-p172">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="d20b8-1037">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-1037">Object</span></span>| <span data-ttu-id="d20b8-1038">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-1039">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d20b8-1040">Objet</span><span class="sxs-lookup"><span data-stu-id="d20b8-1040">Object</span></span>| <span data-ttu-id="d20b8-1041">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-1042">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="d20b8-1043">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d20b8-1043">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="d20b8-1044">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d20b8-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="d20b8-p173">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d20b8-p174">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="d20b8-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d20b8-1049">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="d20b8-1049">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="d20b8-1050">fonction</span><span class="sxs-lookup"><span data-stu-id="d20b8-1050">function</span></span>||<span data-ttu-id="d20b8-1051">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d20b8-1051">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d20b8-1052">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d20b8-1052">Requirements</span></span>

|<span data-ttu-id="d20b8-1053">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d20b8-1053">Requirement</span></span>| <span data-ttu-id="d20b8-1054">Valeur</span><span class="sxs-lookup"><span data-stu-id="d20b8-1054">Value</span></span>|
|---|---|
|[<span data-ttu-id="d20b8-1055">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d20b8-1055">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d20b8-1056">1.2</span><span class="sxs-lookup"><span data-stu-id="d20b8-1056">1.2</span></span>|
|[<span data-ttu-id="d20b8-1057">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d20b8-1057">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d20b8-1058">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d20b8-1058">ReadWriteItem</span></span>|
|[<span data-ttu-id="d20b8-1059">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d20b8-1059">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d20b8-1060">Composition</span><span class="sxs-lookup"><span data-stu-id="d20b8-1060">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d20b8-1061">Exemple</span><span class="sxs-lookup"><span data-stu-id="d20b8-1061">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
