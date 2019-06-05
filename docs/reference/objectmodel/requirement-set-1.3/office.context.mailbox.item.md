---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,3
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 350e476c7f965e825ac08027533b7d0938e1c222
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706364"
---
# <a name="item"></a><span data-ttu-id="946b1-102">élément</span><span class="sxs-lookup"><span data-stu-id="946b1-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="946b1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="946b1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="946b1-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="946b1-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-106">Requirements</span></span>

|<span data-ttu-id="946b1-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-107">Requirement</span></span>| <span data-ttu-id="946b1-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-110">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-110">1.0</span></span>|
|[<span data-ttu-id="946b1-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="946b1-112">Restricted</span></span>|
|[<span data-ttu-id="946b1-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-114">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="946b1-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-115">Example</span></span>

<span data-ttu-id="946b1-116">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="946b1-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="946b1-117">Membres</span><span class="sxs-lookup"><span data-stu-id="946b1-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="946b1-118">pièces jointes: tableau. <[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="946b1-118">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="946b1-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-121">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="946b1-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="946b1-122">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="946b1-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-123">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-123">Type</span></span>

*   <span data-ttu-id="946b1-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="946b1-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-125">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-125">Requirements</span></span>

|<span data-ttu-id="946b1-126">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-126">Requirement</span></span>| <span data-ttu-id="946b1-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-128">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-129">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-129">1.0</span></span>|
|[<span data-ttu-id="946b1-130">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-131">ReadItem</span></span>|
|[<span data-ttu-id="946b1-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-134">Example</span></span>

<span data-ttu-id="946b1-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="946b1-136">CCI: [destinataires](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-136">bcc: [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="946b1-137">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="946b1-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="946b1-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-139">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-139">Type</span></span>

*   [<span data-ttu-id="946b1-140">Destinataires</span><span class="sxs-lookup"><span data-stu-id="946b1-140">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="946b1-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-141">Requirements</span></span>

|<span data-ttu-id="946b1-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-142">Requirement</span></span>| <span data-ttu-id="946b1-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-145">1.1</span><span class="sxs-lookup"><span data-stu-id="946b1-145">1.1</span></span>|
|[<span data-ttu-id="946b1-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-147">ReadItem</span></span>|
|[<span data-ttu-id="946b1-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-149">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="946b1-151">Body: [Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="946b1-151">body: [Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="946b1-152">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-153">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-153">Type</span></span>

*   [<span data-ttu-id="946b1-154">Body</span><span class="sxs-lookup"><span data-stu-id="946b1-154">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="946b1-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-155">Requirements</span></span>

|<span data-ttu-id="946b1-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-156">Requirement</span></span>| <span data-ttu-id="946b1-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-159">1.1</span><span class="sxs-lookup"><span data-stu-id="946b1-159">1.1</span></span>|
|[<span data-ttu-id="946b1-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-161">ReadItem</span></span>|
|[<span data-ttu-id="946b1-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-164">Example</span></span>

<span data-ttu-id="946b1-165">Cet exemple obtient le corps du message en texte brut.</span><span class="sxs-lookup"><span data-stu-id="946b1-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="946b1-166">L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="946b1-167">CC: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-167">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="946b1-168">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="946b1-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="946b1-169">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-170">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-170">Read mode</span></span>

<span data-ttu-id="946b1-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="946b1-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-173">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-173">Compose mode</span></span>

<span data-ttu-id="946b1-174">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="946b1-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="946b1-175">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-175">Type</span></span>

*   <span data-ttu-id="946b1-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-177">Requirements</span></span>

|<span data-ttu-id="946b1-178">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-178">Requirement</span></span>| <span data-ttu-id="946b1-179">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-180">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-181">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-181">1.0</span></span>|
|[<span data-ttu-id="946b1-182">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-183">ReadItem</span></span>|
|[<span data-ttu-id="946b1-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-185">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="946b1-186">(Nullable) conversationId: chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-186">(nullable) conversationId: String</span></span>

<span data-ttu-id="946b1-187">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="946b1-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="946b1-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="946b1-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="946b1-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="946b1-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-192">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-192">Type</span></span>

*   <span data-ttu-id="946b1-193">String</span><span class="sxs-lookup"><span data-stu-id="946b1-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-194">Requirements</span></span>

|<span data-ttu-id="946b1-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-195">Requirement</span></span>| <span data-ttu-id="946b1-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-198">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-198">1.0</span></span>|
|[<span data-ttu-id="946b1-199">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-200">ReadItem</span></span>|
|[<span data-ttu-id="946b1-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="946b1-204">dateTimeCreated: date</span><span class="sxs-lookup"><span data-stu-id="946b1-204">dateTimeCreated: Date</span></span>

<span data-ttu-id="946b1-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-207">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-207">Type</span></span>

*   <span data-ttu-id="946b1-208">Date</span><span class="sxs-lookup"><span data-stu-id="946b1-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-209">Requirements</span></span>

|<span data-ttu-id="946b1-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-210">Requirement</span></span>| <span data-ttu-id="946b1-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-213">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-213">1.0</span></span>|
|[<span data-ttu-id="946b1-214">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-215">ReadItem</span></span>|
|[<span data-ttu-id="946b1-216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-217">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-218">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="946b1-219">dateTimeModified: date</span><span class="sxs-lookup"><span data-stu-id="946b1-219">dateTimeModified: Date</span></span>

<span data-ttu-id="946b1-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-222">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-223">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-223">Type</span></span>

*   <span data-ttu-id="946b1-224">Date</span><span class="sxs-lookup"><span data-stu-id="946b1-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-225">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-225">Requirements</span></span>

|<span data-ttu-id="946b1-226">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-226">Requirement</span></span>| <span data-ttu-id="946b1-227">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-228">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-229">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-229">1.0</span></span>|
|[<span data-ttu-id="946b1-230">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-231">ReadItem</span></span>|
|[<span data-ttu-id="946b1-232">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-233">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-234">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="946b1-235">fin: date | [Fois](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="946b1-235">end: Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="946b1-236">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="946b1-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="946b1-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-239">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-239">Read mode</span></span>

<span data-ttu-id="946b1-240">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="946b1-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-241">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-241">Compose mode</span></span>

<span data-ttu-id="946b1-242">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="946b1-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="946b1-243">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="946b1-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="946b1-244">L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="946b1-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="946b1-245">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-245">Type</span></span>

*   <span data-ttu-id="946b1-246">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="946b1-246">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-247">Requirements</span></span>

|<span data-ttu-id="946b1-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-248">Requirement</span></span>| <span data-ttu-id="946b1-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-251">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-251">1.0</span></span>|
|[<span data-ttu-id="946b1-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-253">ReadItem</span></span>|
|[<span data-ttu-id="946b1-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="946b1-256">de: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="946b1-256">from: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="946b1-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="946b1-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="946b1-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-261">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="946b1-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-262">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-262">Type</span></span>

*   [<span data-ttu-id="946b1-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="946b1-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="946b1-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-264">Requirements</span></span>

|<span data-ttu-id="946b1-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-265">Requirement</span></span>| <span data-ttu-id="946b1-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-268">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-268">1.0</span></span>|
|[<span data-ttu-id="946b1-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-270">ReadItem</span></span>|
|[<span data-ttu-id="946b1-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-272">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-273">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="946b1-274">internetMessageId: chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-274">internetMessageId: String</span></span>

<span data-ttu-id="946b1-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-277">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-277">Type</span></span>

*   <span data-ttu-id="946b1-278">String</span><span class="sxs-lookup"><span data-stu-id="946b1-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-279">Requirements</span></span>

|<span data-ttu-id="946b1-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-280">Requirement</span></span>| <span data-ttu-id="946b1-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-283">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-283">1.0</span></span>|
|[<span data-ttu-id="946b1-284">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-285">ReadItem</span></span>|
|[<span data-ttu-id="946b1-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-287">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="946b1-289">itemClass: chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-289">itemClass: String</span></span>

<span data-ttu-id="946b1-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="946b1-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="946b1-294">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-294">Type</span></span> | <span data-ttu-id="946b1-295">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-295">Description</span></span> | <span data-ttu-id="946b1-296">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="946b1-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="946b1-297">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="946b1-297">Appointment items</span></span> | <span data-ttu-id="946b1-298">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="946b1-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="946b1-299">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="946b1-299">Message items</span></span> | <span data-ttu-id="946b1-300">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="946b1-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="946b1-301">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="946b1-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-302">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-302">Type</span></span>

*   <span data-ttu-id="946b1-303">String</span><span class="sxs-lookup"><span data-stu-id="946b1-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-304">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-304">Requirements</span></span>

|<span data-ttu-id="946b1-305">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-305">Requirement</span></span>| <span data-ttu-id="946b1-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-307">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-308">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-308">1.0</span></span>|
|[<span data-ttu-id="946b1-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-310">ReadItem</span></span>|
|[<span data-ttu-id="946b1-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="946b1-314">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="946b1-314">(nullable) itemId: String</span></span>

<span data-ttu-id="946b1-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-317">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="946b1-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="946b1-318">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="946b1-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="946b1-319">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="946b1-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="946b1-320">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="946b1-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="946b1-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-323">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-323">Type</span></span>

*   <span data-ttu-id="946b1-324">String</span><span class="sxs-lookup"><span data-stu-id="946b1-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-325">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-325">Requirements</span></span>

|<span data-ttu-id="946b1-326">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-326">Requirement</span></span>| <span data-ttu-id="946b1-327">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-328">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-329">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-329">1.0</span></span>|
|[<span data-ttu-id="946b1-330">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-331">ReadItem</span></span>|
|[<span data-ttu-id="946b1-332">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-333">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-334">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-334">Example</span></span>

<span data-ttu-id="946b1-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="946b1-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="946b1-337">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="946b1-337">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="946b1-338">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="946b1-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="946b1-339">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-340">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-340">Type</span></span>

*   [<span data-ttu-id="946b1-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="946b1-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="946b1-342">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-342">Requirements</span></span>

|<span data-ttu-id="946b1-343">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-343">Requirement</span></span>| <span data-ttu-id="946b1-344">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-345">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-346">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-346">1.0</span></span>|
|[<span data-ttu-id="946b1-347">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-348">ReadItem</span></span>|
|[<span data-ttu-id="946b1-349">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-350">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-351">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="946b1-352">Location: String | [Emplacement](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="946b1-352">location: String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="946b1-353">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-354">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-354">Read mode</span></span>

<span data-ttu-id="946b1-355">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-356">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-356">Compose mode</span></span>

<span data-ttu-id="946b1-357">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="946b1-358">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-358">Type</span></span>

*   <span data-ttu-id="946b1-359">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="946b1-359">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-360">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-360">Requirements</span></span>

|<span data-ttu-id="946b1-361">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-361">Requirement</span></span>| <span data-ttu-id="946b1-362">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-363">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-364">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-364">1.0</span></span>|
|[<span data-ttu-id="946b1-365">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-366">ReadItem</span></span>|
|[<span data-ttu-id="946b1-367">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-368">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="946b1-369">normalizedSubject: chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-369">normalizedSubject: String</span></span>

<span data-ttu-id="946b1-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="946b1-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="946b1-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-374">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-374">Type</span></span>

*   <span data-ttu-id="946b1-375">String</span><span class="sxs-lookup"><span data-stu-id="946b1-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-376">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-376">Requirements</span></span>

|<span data-ttu-id="946b1-377">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-377">Requirement</span></span>| <span data-ttu-id="946b1-378">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-379">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-380">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-380">1.0</span></span>|
|[<span data-ttu-id="946b1-381">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-382">ReadItem</span></span>|
|[<span data-ttu-id="946b1-383">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-384">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-385">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="946b1-386">notificationMessages: [notificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="946b1-386">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="946b1-387">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-388">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-388">Type</span></span>

*   [<span data-ttu-id="946b1-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="946b1-389">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="946b1-390">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-390">Requirements</span></span>

|<span data-ttu-id="946b1-391">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-391">Requirement</span></span>| <span data-ttu-id="946b1-392">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-393">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-394">1.3</span><span class="sxs-lookup"><span data-stu-id="946b1-394">1.3</span></span>|
|[<span data-ttu-id="946b1-395">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-396">ReadItem</span></span>|
|[<span data-ttu-id="946b1-397">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-398">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-399">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="946b1-400">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[](/javascript/api/outlook_1_3/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="946b1-400">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="946b1-401">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="946b1-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="946b1-402">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-403">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-403">Read mode</span></span>

<span data-ttu-id="946b1-404">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="946b1-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-405">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-405">Compose mode</span></span>

<span data-ttu-id="946b1-406">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="946b1-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="946b1-407">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-407">Type</span></span>

*   <span data-ttu-id="946b1-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-409">Requirements</span></span>

|<span data-ttu-id="946b1-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-410">Requirement</span></span>| <span data-ttu-id="946b1-411">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-412">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-413">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-413">1.0</span></span>|
|[<span data-ttu-id="946b1-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-415">ReadItem</span></span>|
|[<span data-ttu-id="946b1-416">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-417">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="946b1-418">Organisateur: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="946b1-418">organizer: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="946b1-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-421">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-421">Type</span></span>

*   [<span data-ttu-id="946b1-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="946b1-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="946b1-423">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-423">Requirements</span></span>

|<span data-ttu-id="946b1-424">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-424">Requirement</span></span>| <span data-ttu-id="946b1-425">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-426">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-427">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-427">1.0</span></span>|
|[<span data-ttu-id="946b1-428">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-429">ReadItem</span></span>|
|[<span data-ttu-id="946b1-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-431">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-432">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="946b1-433">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[](/javascript/api/outlook_1_3/office.recipients) des destinataires de tableau. <</span><span class="sxs-lookup"><span data-stu-id="946b1-433">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="946b1-434">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="946b1-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="946b1-435">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-436">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-436">Read mode</span></span>

<span data-ttu-id="946b1-437">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="946b1-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-438">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-438">Compose mode</span></span>

<span data-ttu-id="946b1-439">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="946b1-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="946b1-440">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-440">Type</span></span>

*   <span data-ttu-id="946b1-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-442">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-442">Requirements</span></span>

|<span data-ttu-id="946b1-443">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-443">Requirement</span></span>| <span data-ttu-id="946b1-444">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-445">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-446">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-446">1.0</span></span>|
|[<span data-ttu-id="946b1-447">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-448">ReadItem</span></span>|
|[<span data-ttu-id="946b1-449">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-450">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="946b1-451">expéditeur: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="946b1-451">sender: [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="946b1-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="946b1-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="946b1-p127">Les propriétés [`from`](#from-emailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="946b1-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-456">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="946b1-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="946b1-457">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-457">Type</span></span>

*   [<span data-ttu-id="946b1-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="946b1-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="946b1-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-459">Requirements</span></span>

|<span data-ttu-id="946b1-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-460">Requirement</span></span>| <span data-ttu-id="946b1-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-463">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-463">1.0</span></span>|
|[<span data-ttu-id="946b1-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-465">ReadItem</span></span>|
|[<span data-ttu-id="946b1-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-467">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="946b1-469">début: date | [Fois](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="946b1-469">start: Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="946b1-470">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="946b1-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="946b1-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-473">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-473">Read mode</span></span>

<span data-ttu-id="946b1-474">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="946b1-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-475">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-475">Compose mode</span></span>

<span data-ttu-id="946b1-476">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="946b1-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="946b1-477">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="946b1-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="946b1-478">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="946b1-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="946b1-479">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-479">Type</span></span>

*   <span data-ttu-id="946b1-480">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="946b1-480">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-481">Requirements</span></span>

|<span data-ttu-id="946b1-482">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-482">Requirement</span></span>| <span data-ttu-id="946b1-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-485">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-485">1.0</span></span>|
|[<span data-ttu-id="946b1-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-487">ReadItem</span></span>|
|[<span data-ttu-id="946b1-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-489">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-489">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="946b1-490">Subject: String | [Objet](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="946b1-490">subject: String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="946b1-491">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="946b1-492">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="946b1-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-493">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-493">Read mode</span></span>

<span data-ttu-id="946b1-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="946b1-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-496">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-496">Compose mode</span></span>

<span data-ttu-id="946b1-497">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="946b1-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="946b1-498">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-498">Type</span></span>

*   <span data-ttu-id="946b1-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="946b1-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-500">Requirements</span></span>

|<span data-ttu-id="946b1-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-501">Requirement</span></span>| <span data-ttu-id="946b1-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-504">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-504">1.0</span></span>|
|[<span data-ttu-id="946b1-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-506">ReadItem</span></span>|
|[<span data-ttu-id="946b1-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-508">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-508">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="946b1-509">to: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-509">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="946b1-510">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="946b1-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="946b1-511">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="946b1-512">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-512">Read mode</span></span>

<span data-ttu-id="946b1-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="946b1-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="946b1-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="946b1-515">Compose mode</span></span>

<span data-ttu-id="946b1-516">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="946b1-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="946b1-517">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-517">Type</span></span>

*   <span data-ttu-id="946b1-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="946b1-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-519">Requirements</span></span>

|<span data-ttu-id="946b1-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-520">Requirement</span></span>| <span data-ttu-id="946b1-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-523">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-523">1.0</span></span>|
|[<span data-ttu-id="946b1-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-525">ReadItem</span></span>|
|[<span data-ttu-id="946b1-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-527">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="946b1-528">Méthodes</span><span class="sxs-lookup"><span data-stu-id="946b1-528">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="946b1-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="946b1-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="946b1-530">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="946b1-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="946b1-531">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="946b1-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="946b1-532">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="946b1-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-533">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-533">Parameters</span></span>

|<span data-ttu-id="946b1-534">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-534">Name</span></span>| <span data-ttu-id="946b1-535">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-535">Type</span></span>| <span data-ttu-id="946b1-536">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-536">Attributes</span></span>| <span data-ttu-id="946b1-537">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="946b1-538">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-538">String</span></span>||<span data-ttu-id="946b1-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="946b1-541">String</span><span class="sxs-lookup"><span data-stu-id="946b1-541">String</span></span>||<span data-ttu-id="946b1-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="946b1-544">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-544">Object</span></span>| <span data-ttu-id="946b1-545">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-545">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-546">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="946b1-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="946b1-547">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-547">Object</span></span>| <span data-ttu-id="946b1-548">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-548">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-549">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="946b1-550">fonction</span><span class="sxs-lookup"><span data-stu-id="946b1-550">function</span></span>| <span data-ttu-id="946b1-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-551">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-552">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="946b1-553">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="946b1-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="946b1-554">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="946b1-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="946b1-555">Erreurs</span><span class="sxs-lookup"><span data-stu-id="946b1-555">Errors</span></span>

| <span data-ttu-id="946b1-556">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="946b1-556">Error code</span></span> | <span data-ttu-id="946b1-557">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="946b1-558">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="946b1-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="946b1-559">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="946b1-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="946b1-560">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="946b1-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="946b1-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-561">Requirements</span></span>

|<span data-ttu-id="946b1-562">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-562">Requirement</span></span>| <span data-ttu-id="946b1-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-565">1.1</span><span class="sxs-lookup"><span data-stu-id="946b1-565">1.1</span></span>|
|[<span data-ttu-id="946b1-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="946b1-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="946b1-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-569">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-570">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="946b1-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="946b1-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="946b1-572">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="946b1-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="946b1-576">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="946b1-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="946b1-577">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="946b1-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-578">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-578">Parameters</span></span>

|<span data-ttu-id="946b1-579">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-579">Name</span></span>| <span data-ttu-id="946b1-580">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-580">Type</span></span>| <span data-ttu-id="946b1-581">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-581">Attributes</span></span>| <span data-ttu-id="946b1-582">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="946b1-583">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-583">String</span></span>||<span data-ttu-id="946b1-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="946b1-586">String</span><span class="sxs-lookup"><span data-stu-id="946b1-586">String</span></span>||<span data-ttu-id="946b1-587">Objet de l’élément à joindre.</span><span class="sxs-lookup"><span data-stu-id="946b1-587">The subject of the item to be attached.</span></span> <span data-ttu-id="946b1-588">La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="946b1-589">Object</span><span class="sxs-lookup"><span data-stu-id="946b1-589">Object</span></span>| <span data-ttu-id="946b1-590">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-590">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-591">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="946b1-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="946b1-592">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-592">Object</span></span>| <span data-ttu-id="946b1-593">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-593">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-594">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="946b1-595">fonction</span><span class="sxs-lookup"><span data-stu-id="946b1-595">function</span></span>| <span data-ttu-id="946b1-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-596">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-597">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="946b1-598">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="946b1-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="946b1-599">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="946b1-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="946b1-600">Erreurs</span><span class="sxs-lookup"><span data-stu-id="946b1-600">Errors</span></span>

| <span data-ttu-id="946b1-601">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="946b1-601">Error code</span></span> | <span data-ttu-id="946b1-602">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="946b1-603">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="946b1-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="946b1-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-604">Requirements</span></span>

|<span data-ttu-id="946b1-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-605">Requirement</span></span>| <span data-ttu-id="946b1-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-607">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-608">1.1</span><span class="sxs-lookup"><span data-stu-id="946b1-608">1.1</span></span>|
|[<span data-ttu-id="946b1-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="946b1-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="946b1-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-612">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-613">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-613">Example</span></span>

<span data-ttu-id="946b1-614">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="946b1-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="946b1-615">close()</span><span class="sxs-lookup"><span data-stu-id="946b1-615">close()</span></span>

<span data-ttu-id="946b1-616">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="946b1-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="946b1-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="946b1-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-619">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="946b1-620">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="946b1-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-621">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-621">Requirements</span></span>

|<span data-ttu-id="946b1-622">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-622">Requirement</span></span>| <span data-ttu-id="946b1-623">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-624">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-625">1.3</span><span class="sxs-lookup"><span data-stu-id="946b1-625">1.3</span></span>|
|[<span data-ttu-id="946b1-626">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-627">Restreinte</span><span class="sxs-lookup"><span data-stu-id="946b1-627">Restricted</span></span>|
|[<span data-ttu-id="946b1-628">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-629">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="946b1-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="946b1-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="946b1-631">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="946b1-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-632">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="946b1-633">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="946b1-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="946b1-634">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="946b1-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="946b1-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="946b1-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-638">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-638">Parameters</span></span>

|<span data-ttu-id="946b1-639">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-639">Name</span></span>| <span data-ttu-id="946b1-640">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-640">Type</span></span>| <span data-ttu-id="946b1-641">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="946b1-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="946b1-642">String &#124; Object</span></span>| |<span data-ttu-id="946b1-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="946b1-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="946b1-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="946b1-645">**OR**</span></span><br/><span data-ttu-id="946b1-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="946b1-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="946b1-648">String</span><span class="sxs-lookup"><span data-stu-id="946b1-648">String</span></span> | <span data-ttu-id="946b1-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-649">&lt;optional&gt;</span></span> | <span data-ttu-id="946b1-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="946b1-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="946b1-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="946b1-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-653">&lt;optional&gt;</span></span> | <span data-ttu-id="946b1-654">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="946b1-655">String</span><span class="sxs-lookup"><span data-stu-id="946b1-655">String</span></span> | | <span data-ttu-id="946b1-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="946b1-658">String</span><span class="sxs-lookup"><span data-stu-id="946b1-658">String</span></span> | | <span data-ttu-id="946b1-659">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="946b1-660">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-660">String</span></span> | | <span data-ttu-id="946b1-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="946b1-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="946b1-663">String</span><span class="sxs-lookup"><span data-stu-id="946b1-663">String</span></span> | | <span data-ttu-id="946b1-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="946b1-667">function</span><span class="sxs-lookup"><span data-stu-id="946b1-667">function</span></span> | <span data-ttu-id="946b1-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-668">&lt;optional&gt;</span></span> | <span data-ttu-id="946b1-669">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="946b1-670">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-670">Requirements</span></span>

|<span data-ttu-id="946b1-671">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-671">Requirement</span></span>| <span data-ttu-id="946b1-672">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-673">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-674">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-674">1.0</span></span>|
|[<span data-ttu-id="946b1-675">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-676">ReadItem</span></span>|
|[<span data-ttu-id="946b1-677">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-678">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="946b1-679">Exemples</span><span class="sxs-lookup"><span data-stu-id="946b1-679">Examples</span></span>

<span data-ttu-id="946b1-680">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="946b1-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="946b1-681">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="946b1-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="946b1-682">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="946b1-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="946b1-683">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="946b1-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="946b1-684">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="946b1-685">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="946b1-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="946b1-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="946b1-687">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="946b1-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-688">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="946b1-689">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="946b1-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="946b1-690">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="946b1-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="946b1-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="946b1-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-694">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-694">Parameters</span></span>

|<span data-ttu-id="946b1-695">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-695">Name</span></span>| <span data-ttu-id="946b1-696">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-696">Type</span></span>| <span data-ttu-id="946b1-697">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="946b1-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="946b1-698">String &#124; Object</span></span>| | <span data-ttu-id="946b1-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="946b1-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="946b1-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="946b1-701">**OR**</span></span><br/><span data-ttu-id="946b1-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="946b1-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="946b1-704">String</span><span class="sxs-lookup"><span data-stu-id="946b1-704">String</span></span> | <span data-ttu-id="946b1-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-705">&lt;optional&gt;</span></span> | <span data-ttu-id="946b1-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="946b1-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="946b1-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="946b1-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-709">&lt;optional&gt;</span></span> | <span data-ttu-id="946b1-710">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="946b1-711">String</span><span class="sxs-lookup"><span data-stu-id="946b1-711">String</span></span> | | <span data-ttu-id="946b1-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="946b1-714">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-714">String</span></span> | | <span data-ttu-id="946b1-715">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="946b1-716">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-716">String</span></span> | | <span data-ttu-id="946b1-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="946b1-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="946b1-719">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-719">String</span></span> | | <span data-ttu-id="946b1-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="946b1-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="946b1-723">function</span><span class="sxs-lookup"><span data-stu-id="946b1-723">function</span></span> | <span data-ttu-id="946b1-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-724">&lt;optional&gt;</span></span> | <span data-ttu-id="946b1-725">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="946b1-726">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-726">Requirements</span></span>

|<span data-ttu-id="946b1-727">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-727">Requirement</span></span>| <span data-ttu-id="946b1-728">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-729">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-730">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-730">1.0</span></span>|
|[<span data-ttu-id="946b1-731">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-732">ReadItem</span></span>|
|[<span data-ttu-id="946b1-733">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-734">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="946b1-735">Exemples</span><span class="sxs-lookup"><span data-stu-id="946b1-735">Examples</span></span>

<span data-ttu-id="946b1-736">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="946b1-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="946b1-737">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="946b1-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="946b1-738">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="946b1-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="946b1-739">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="946b1-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="946b1-740">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="946b1-741">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="946b1-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="946b1-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="946b1-743">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="946b1-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-744">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-745">Requirements</span></span>

|<span data-ttu-id="946b1-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-746">Requirement</span></span>| <span data-ttu-id="946b1-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-749">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-749">1.0</span></span>|
|[<span data-ttu-id="946b1-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-751">ReadItem</span></span>|
|[<span data-ttu-id="946b1-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-753">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="946b1-754">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="946b1-754">Returns:</span></span>

<span data-ttu-id="946b1-755">Type : [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="946b1-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="946b1-756">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-756">Example</span></span>

<span data-ttu-id="946b1-757">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="946b1-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="946b1-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="946b1-759">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="946b1-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-760">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-761">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-761">Parameters</span></span>

|<span data-ttu-id="946b1-762">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-762">Name</span></span>| <span data-ttu-id="946b1-763">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-763">Type</span></span>| <span data-ttu-id="946b1-764">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="946b1-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="946b1-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="946b1-766">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="946b1-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="946b1-767">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-767">Requirements</span></span>

|<span data-ttu-id="946b1-768">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-768">Requirement</span></span>| <span data-ttu-id="946b1-769">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-770">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-771">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-771">1.0</span></span>|
|[<span data-ttu-id="946b1-772">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-773">Restreinte</span><span class="sxs-lookup"><span data-stu-id="946b1-773">Restricted</span></span>|
|[<span data-ttu-id="946b1-774">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-775">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="946b1-776">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="946b1-776">Returns:</span></span>

<span data-ttu-id="946b1-777">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="946b1-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="946b1-778">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="946b1-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="946b1-779">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="946b1-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="946b1-780">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="946b1-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="946b1-781">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="946b1-781">Value of `entityType`</span></span> | <span data-ttu-id="946b1-782">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="946b1-782">Type of objects in returned array</span></span> | <span data-ttu-id="946b1-783">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="946b1-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="946b1-784">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-784">String</span></span> | <span data-ttu-id="946b1-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="946b1-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="946b1-786">Contact</span><span class="sxs-lookup"><span data-stu-id="946b1-786">Contact</span></span> | <span data-ttu-id="946b1-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="946b1-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="946b1-788">String</span><span class="sxs-lookup"><span data-stu-id="946b1-788">String</span></span> | <span data-ttu-id="946b1-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="946b1-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="946b1-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="946b1-790">MeetingSuggestion</span></span> | <span data-ttu-id="946b1-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="946b1-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="946b1-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="946b1-792">PhoneNumber</span></span> | <span data-ttu-id="946b1-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="946b1-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="946b1-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="946b1-794">TaskSuggestion</span></span> | <span data-ttu-id="946b1-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="946b1-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="946b1-796">String</span><span class="sxs-lookup"><span data-stu-id="946b1-796">String</span></span> | <span data-ttu-id="946b1-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="946b1-797">**Restricted**</span></span> |

<span data-ttu-id="946b1-798">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="946b1-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="946b1-799">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-799">Example</span></span>

<span data-ttu-id="946b1-800">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="946b1-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="946b1-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="946b1-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="946b1-802">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="946b1-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-803">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="946b1-804">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="946b1-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-805">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-805">Parameters</span></span>

|<span data-ttu-id="946b1-806">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-806">Name</span></span>| <span data-ttu-id="946b1-807">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-807">Type</span></span>| <span data-ttu-id="946b1-808">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="946b1-809">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-809">String</span></span>|<span data-ttu-id="946b1-810">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="946b1-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="946b1-811">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-811">Requirements</span></span>

|<span data-ttu-id="946b1-812">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-812">Requirement</span></span>| <span data-ttu-id="946b1-813">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-814">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-815">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-815">1.0</span></span>|
|[<span data-ttu-id="946b1-816">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-817">ReadItem</span></span>|
|[<span data-ttu-id="946b1-818">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-819">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="946b1-820">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="946b1-820">Returns:</span></span>

<span data-ttu-id="946b1-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="946b1-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="946b1-823">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="946b1-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="946b1-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="946b1-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="946b1-825">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="946b1-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-826">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="946b1-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="946b1-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="946b1-830">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="946b1-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="946b1-831">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="946b1-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="946b1-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="946b1-835">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-835">Requirements</span></span>

|<span data-ttu-id="946b1-836">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-836">Requirement</span></span>| <span data-ttu-id="946b1-837">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-838">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-839">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-839">1.0</span></span>|
|[<span data-ttu-id="946b1-840">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-841">ReadItem</span></span>|
|[<span data-ttu-id="946b1-842">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-843">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="946b1-844">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="946b1-844">Returns:</span></span>

<span data-ttu-id="946b1-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="946b1-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="946b1-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="946b1-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="946b1-848">Object</span><span class="sxs-lookup"><span data-stu-id="946b1-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="946b1-849">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-849">Example</span></span>

<span data-ttu-id="946b1-850">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="946b1-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="946b1-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="946b1-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="946b1-852">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="946b1-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-853">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="946b1-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="946b1-854">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="946b1-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="946b1-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="946b1-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-857">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-857">Parameters</span></span>

|<span data-ttu-id="946b1-858">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-858">Name</span></span>| <span data-ttu-id="946b1-859">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-859">Type</span></span>| <span data-ttu-id="946b1-860">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="946b1-861">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-861">String</span></span>|<span data-ttu-id="946b1-862">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="946b1-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="946b1-863">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-863">Requirements</span></span>

|<span data-ttu-id="946b1-864">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-864">Requirement</span></span>| <span data-ttu-id="946b1-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-866">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-867">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-867">1.0</span></span>|
|[<span data-ttu-id="946b1-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-869">ReadItem</span></span>|
|[<span data-ttu-id="946b1-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="946b1-872">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="946b1-872">Returns:</span></span>

<span data-ttu-id="946b1-873">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="946b1-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="946b1-874">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="946b1-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="946b1-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="946b1-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="946b1-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="946b1-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="946b1-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="946b1-878">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="946b1-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="946b1-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="946b1-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-881">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-881">Parameters</span></span>

|<span data-ttu-id="946b1-882">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-882">Name</span></span>| <span data-ttu-id="946b1-883">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-883">Type</span></span>| <span data-ttu-id="946b1-884">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-884">Attributes</span></span>| <span data-ttu-id="946b1-885">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="946b1-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="946b1-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="946b1-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="946b1-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="946b1-890">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-890">Object</span></span>| <span data-ttu-id="946b1-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-891">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-892">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="946b1-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="946b1-893">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-893">Object</span></span>| <span data-ttu-id="946b1-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-894">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-895">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="946b1-896">fonction</span><span class="sxs-lookup"><span data-stu-id="946b1-896">function</span></span>||<span data-ttu-id="946b1-897">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="946b1-898">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="946b1-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="946b1-899">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="946b1-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="946b1-900">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-900">Requirements</span></span>

|<span data-ttu-id="946b1-901">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-901">Requirement</span></span>| <span data-ttu-id="946b1-902">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-903">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-904">1.2</span><span class="sxs-lookup"><span data-stu-id="946b1-904">1.2</span></span>|
|[<span data-ttu-id="946b1-905">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="946b1-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="946b1-907">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-908">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="946b1-909">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="946b1-909">Returns:</span></span>

<span data-ttu-id="946b1-910">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="946b1-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="946b1-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="946b1-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="946b1-912">String</span><span class="sxs-lookup"><span data-stu-id="946b1-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="946b1-913">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-913">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="946b1-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="946b1-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="946b1-915">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="946b1-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="946b1-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="946b1-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-919">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-919">Parameters</span></span>

|<span data-ttu-id="946b1-920">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-920">Name</span></span>| <span data-ttu-id="946b1-921">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-921">Type</span></span>| <span data-ttu-id="946b1-922">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-922">Attributes</span></span>| <span data-ttu-id="946b1-923">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="946b1-924">function</span><span class="sxs-lookup"><span data-stu-id="946b1-924">function</span></span>||<span data-ttu-id="946b1-925">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="946b1-926">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="946b1-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="946b1-927">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="946b1-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="946b1-928">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-928">Object</span></span>| <span data-ttu-id="946b1-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-929">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-930">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="946b1-931">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="946b1-932">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-932">Requirements</span></span>

|<span data-ttu-id="946b1-933">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-933">Requirement</span></span>| <span data-ttu-id="946b1-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-935">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-936">1.0</span><span class="sxs-lookup"><span data-stu-id="946b1-936">1.0</span></span>|
|[<span data-ttu-id="946b1-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="946b1-938">ReadItem</span></span>|
|[<span data-ttu-id="946b1-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-940">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="946b1-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-941">Example</span></span>

<span data-ttu-id="946b1-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="946b1-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="946b1-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="946b1-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="946b1-946">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="946b1-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="946b1-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="946b1-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-951">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-951">Parameters</span></span>

|<span data-ttu-id="946b1-952">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-952">Name</span></span>| <span data-ttu-id="946b1-953">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-953">Type</span></span>| <span data-ttu-id="946b1-954">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-954">Attributes</span></span>| <span data-ttu-id="946b1-955">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="946b1-956">Chaîne</span><span class="sxs-lookup"><span data-stu-id="946b1-956">String</span></span>||<span data-ttu-id="946b1-957">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="946b1-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="946b1-958">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-958">Object</span></span>| <span data-ttu-id="946b1-959">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-959">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-960">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="946b1-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="946b1-961">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-961">Object</span></span>| <span data-ttu-id="946b1-962">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-962">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-963">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="946b1-964">fonction</span><span class="sxs-lookup"><span data-stu-id="946b1-964">function</span></span>| <span data-ttu-id="946b1-965">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-965">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-966">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="946b1-967">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="946b1-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="946b1-968">Erreurs</span><span class="sxs-lookup"><span data-stu-id="946b1-968">Errors</span></span>

| <span data-ttu-id="946b1-969">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="946b1-969">Error code</span></span> | <span data-ttu-id="946b1-970">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="946b1-971">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="946b1-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="946b1-972">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-972">Requirements</span></span>

|<span data-ttu-id="946b1-973">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-973">Requirement</span></span>| <span data-ttu-id="946b1-974">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-975">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-976">1.1</span><span class="sxs-lookup"><span data-stu-id="946b1-976">1.1</span></span>|
|[<span data-ttu-id="946b1-977">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="946b1-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="946b1-979">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-980">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-981">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-981">Example</span></span>

<span data-ttu-id="946b1-982">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="946b1-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="946b1-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="946b1-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="946b1-984">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="946b1-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="946b1-p166">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="946b1-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-988">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="946b1-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="946b1-989">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="946b1-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="946b1-p168">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="946b1-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="946b1-993">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="946b1-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="946b1-994">Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="946b1-994">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="946b1-995">La `saveAsync` méthode échoue lorsqu’elle est appelée à partir d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="946b1-995">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="946b1-996">Consultez la rubrique [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide de l’API Office js](https://support.microsoft.com/help/4505745) pour obtenir une solution de contournement.</span><span class="sxs-lookup"><span data-stu-id="946b1-996">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="946b1-997">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="946b1-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-998">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-998">Parameters</span></span>

|<span data-ttu-id="946b1-999">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-999">Name</span></span>| <span data-ttu-id="946b1-1000">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-1000">Type</span></span>| <span data-ttu-id="946b1-1001">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-1001">Attributes</span></span>| <span data-ttu-id="946b1-1002">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="946b1-1003">Object</span><span class="sxs-lookup"><span data-stu-id="946b1-1003">Object</span></span>| <span data-ttu-id="946b1-1004">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-1005">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="946b1-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="946b1-1006">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-1006">Object</span></span>| <span data-ttu-id="946b1-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-1008">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="946b1-1009">fonction</span><span class="sxs-lookup"><span data-stu-id="946b1-1009">function</span></span>||<span data-ttu-id="946b1-1010">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="946b1-1011">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="946b1-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="946b1-1012">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-1012">Requirements</span></span>

|<span data-ttu-id="946b1-1013">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-1013">Requirement</span></span>| <span data-ttu-id="946b1-1014">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-1015">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="946b1-1016">1.3</span></span>|
|[<span data-ttu-id="946b1-1017">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-1017">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="946b1-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="946b1-1019">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-1019">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-1020">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="946b1-1021">範例</span><span class="sxs-lookup"><span data-stu-id="946b1-1021">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="946b1-p170">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="946b1-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="946b1-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="946b1-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="946b1-1025">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="946b1-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="946b1-p171">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="946b1-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="946b1-1029">Paramètres</span><span class="sxs-lookup"><span data-stu-id="946b1-1029">Parameters</span></span>

|<span data-ttu-id="946b1-1030">Nom</span><span class="sxs-lookup"><span data-stu-id="946b1-1030">Name</span></span>| <span data-ttu-id="946b1-1031">Type</span><span class="sxs-lookup"><span data-stu-id="946b1-1031">Type</span></span>| <span data-ttu-id="946b1-1032">Attributs</span><span class="sxs-lookup"><span data-stu-id="946b1-1032">Attributes</span></span>| <span data-ttu-id="946b1-1033">Description</span><span class="sxs-lookup"><span data-stu-id="946b1-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="946b1-1034">String</span><span class="sxs-lookup"><span data-stu-id="946b1-1034">String</span></span>||<span data-ttu-id="946b1-p172">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="946b1-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="946b1-1038">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-1038">Object</span></span>| <span data-ttu-id="946b1-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-1040">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="946b1-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="946b1-1041">Objet</span><span class="sxs-lookup"><span data-stu-id="946b1-1041">Object</span></span>| <span data-ttu-id="946b1-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-1043">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="946b1-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="946b1-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="946b1-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="946b1-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="946b1-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="946b1-p173">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="946b1-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="946b1-p174">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="946b1-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="946b1-1050">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="946b1-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="946b1-1051">fonction</span><span class="sxs-lookup"><span data-stu-id="946b1-1051">function</span></span>||<span data-ttu-id="946b1-1052">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="946b1-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="946b1-1053">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="946b1-1053">Requirements</span></span>

|<span data-ttu-id="946b1-1054">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="946b1-1054">Requirement</span></span>| <span data-ttu-id="946b1-1055">Valeur</span><span class="sxs-lookup"><span data-stu-id="946b1-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="946b1-1056">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="946b1-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="946b1-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="946b1-1057">1.2</span></span>|
|[<span data-ttu-id="946b1-1058">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="946b1-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="946b1-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="946b1-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="946b1-1060">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="946b1-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="946b1-1061">Composition</span><span class="sxs-lookup"><span data-stu-id="946b1-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="946b1-1062">Exemple</span><span class="sxs-lookup"><span data-stu-id="946b1-1062">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
