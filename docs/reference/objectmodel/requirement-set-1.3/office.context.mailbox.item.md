---
title: Office.Context.Mailbox.Item - exigence défini 1.3
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: c418c69e369e5f8ed6da151345013897f1a87e26
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387323"
---
# <a name="item"></a><span data-ttu-id="b722d-102">élément</span><span class="sxs-lookup"><span data-stu-id="b722d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b722d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b722d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b722d-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="b722d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-106">Requirements</span></span>

|<span data-ttu-id="b722d-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-107">Requirement</span></span>| <span data-ttu-id="b722d-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-110">1.0</span></span>|
|[<span data-ttu-id="b722d-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-112">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b722d-112">Restricted</span></span>|
|[<span data-ttu-id="b722d-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="b722d-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-115">Example</span></span>

<span data-ttu-id="b722d-116">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="b722d-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b722d-117">Membres</span><span class="sxs-lookup"><span data-stu-id="b722d-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="b722d-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b722d-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="b722d-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-121">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="b722d-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b722d-122">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b722d-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-123">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-123">Type:</span></span>

*   <span data-ttu-id="b722d-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b722d-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-125">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-125">Requirements</span></span>

|<span data-ttu-id="b722d-126">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-126">Requirement</span></span>| <span data-ttu-id="b722d-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-128">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-129">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-129">1.0</span></span>|
|[<span data-ttu-id="b722d-130">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-131">ReadItem</span></span>|
|[<span data-ttu-id="b722d-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-134">Example</span></span>

<span data-ttu-id="b722d-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b722d-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b722d-137">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="b722d-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b722d-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-139">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-139">Type:</span></span>

*   [<span data-ttu-id="b722d-140">Destinataires</span><span class="sxs-lookup"><span data-stu-id="b722d-140">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b722d-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-141">Requirements</span></span>

|<span data-ttu-id="b722d-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-142">Requirement</span></span>| <span data-ttu-id="b722d-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-145">1.1</span><span class="sxs-lookup"><span data-stu-id="b722d-145">1.1</span></span>|
|[<span data-ttu-id="b722d-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-147">ReadItem</span></span>|
|[<span data-ttu-id="b722d-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-149">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-150">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="b722d-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="b722d-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="b722d-152">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-153">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-153">Type:</span></span>

*   [<span data-ttu-id="b722d-154">Corps</span><span class="sxs-lookup"><span data-stu-id="b722d-154">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="b722d-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-155">Requirements</span></span>

|<span data-ttu-id="b722d-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-156">Requirement</span></span>| <span data-ttu-id="b722d-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b722d-159">1.1</span></span>|
|[<span data-ttu-id="b722d-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-161">ReadItem</span></span>|
|[<span data-ttu-id="b722d-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b722d-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b722d-165">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="b722d-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b722d-166">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-167">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-167">Read mode</span></span>

<span data-ttu-id="b722d-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="b722d-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-170">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-170">Compose mode</span></span>

<span data-ttu-id="b722d-171">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="b722d-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-172">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-172">Type:</span></span>

*   <span data-ttu-id="b722d-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-174">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-174">Requirements</span></span>

|<span data-ttu-id="b722d-175">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-175">Requirement</span></span>| <span data-ttu-id="b722d-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-177">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-178">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-178">1.0</span></span>|
|[<span data-ttu-id="b722d-179">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-180">ReadItem</span></span>|
|[<span data-ttu-id="b722d-181">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-182">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-183">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-183">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b722d-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b722d-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="b722d-185">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="b722d-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b722d-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="b722d-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b722d-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="b722d-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-190">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-190">Type:</span></span>

*   <span data-ttu-id="b722d-191">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-192">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-192">Requirements</span></span>

|<span data-ttu-id="b722d-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-193">Requirement</span></span>| <span data-ttu-id="b722d-194">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-195">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-196">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-196">1.0</span></span>|
|[<span data-ttu-id="b722d-197">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-198">ReadItem</span></span>|
|[<span data-ttu-id="b722d-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b722d-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b722d-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="b722d-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-204">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-204">Type:</span></span>

*   <span data-ttu-id="b722d-205">Date</span><span class="sxs-lookup"><span data-stu-id="b722d-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-206">Requirements</span></span>

|<span data-ttu-id="b722d-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-207">Requirement</span></span>| <span data-ttu-id="b722d-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-210">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-210">1.0</span></span>|
|[<span data-ttu-id="b722d-211">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-212">ReadItem</span></span>|
|[<span data-ttu-id="b722d-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-214">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-215">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b722d-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b722d-216">dateTimeModified :Date</span></span>

<span data-ttu-id="b722d-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-219">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-220">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-220">Type:</span></span>

*   <span data-ttu-id="b722d-221">Date</span><span class="sxs-lookup"><span data-stu-id="b722d-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-222">Requirements</span></span>

|<span data-ttu-id="b722d-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-223">Requirement</span></span>| <span data-ttu-id="b722d-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-226">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-226">1.0</span></span>|
|[<span data-ttu-id="b722d-227">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-228">ReadItem</span></span>|
|[<span data-ttu-id="b722d-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-230">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-231">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-231">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="b722d-232">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b722d-232">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="b722d-233">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b722d-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="b722d-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-236">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-236">Read mode</span></span>

<span data-ttu-id="b722d-237">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="b722d-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-238">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-238">Compose mode</span></span>

<span data-ttu-id="b722d-239">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b722d-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b722d-240">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="b722d-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-241">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-241">Type:</span></span>

*   <span data-ttu-id="b722d-242">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b722d-242">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-243">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-243">Requirements</span></span>

|<span data-ttu-id="b722d-244">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-244">Requirement</span></span>| <span data-ttu-id="b722d-245">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-246">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-247">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-247">1.0</span></span>|
|[<span data-ttu-id="b722d-248">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-249">ReadItem</span></span>|
|[<span data-ttu-id="b722d-250">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-251">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-252">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-252">Example</span></span>

<span data-ttu-id="b722d-253">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b722d-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="b722d-254">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b722d-254">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="b722d-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b722d-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="b722d-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-259">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b722d-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-260">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-260">Type:</span></span>

*   [<span data-ttu-id="b722d-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b722d-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b722d-262">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-262">Requirements</span></span>

|<span data-ttu-id="b722d-263">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-263">Requirement</span></span>| <span data-ttu-id="b722d-264">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-265">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-266">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-266">1.0</span></span>|
|[<span data-ttu-id="b722d-267">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-268">ReadItem</span></span>|
|[<span data-ttu-id="b722d-269">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-270">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b722d-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b722d-271">internetMessageId :String</span></span>

<span data-ttu-id="b722d-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-274">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-274">Type:</span></span>

*   <span data-ttu-id="b722d-275">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-276">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-276">Requirements</span></span>

|<span data-ttu-id="b722d-277">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-277">Requirement</span></span>| <span data-ttu-id="b722d-278">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-279">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-280">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-280">1.0</span></span>|
|[<span data-ttu-id="b722d-281">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-282">ReadItem</span></span>|
|[<span data-ttu-id="b722d-283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-284">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-285">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-285">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b722d-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b722d-286">itemClass :String</span></span>

<span data-ttu-id="b722d-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b722d-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b722d-291">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-291">Type</span></span> | <span data-ttu-id="b722d-292">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-292">Description</span></span> | <span data-ttu-id="b722d-293">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="b722d-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b722d-294">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="b722d-294">Appointment items</span></span> | <span data-ttu-id="b722d-295">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="b722d-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="b722d-296">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="b722d-296">Message items</span></span> | <span data-ttu-id="b722d-297">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="b722d-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b722d-298">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="b722d-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-299">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-299">Type:</span></span>

*   <span data-ttu-id="b722d-300">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-301">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-301">Requirements</span></span>

|<span data-ttu-id="b722d-302">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-302">Requirement</span></span>| <span data-ttu-id="b722d-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-304">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-305">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-305">1.0</span></span>|
|[<span data-ttu-id="b722d-306">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-307">ReadItem</span></span>|
|[<span data-ttu-id="b722d-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-310">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b722d-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b722d-311">(nullable) itemId :String</span></span>

<span data-ttu-id="b722d-p117">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-314">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="b722d-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b722d-315">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="b722d-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b722d-316">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="b722d-316">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b722d-317">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b722d-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b722d-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-320">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-320">Type:</span></span>

*   <span data-ttu-id="b722d-321">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-321">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-322">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-322">Requirements</span></span>

|<span data-ttu-id="b722d-323">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-323">Requirement</span></span>| <span data-ttu-id="b722d-324">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-325">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-326">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-326">1.0</span></span>|
|[<span data-ttu-id="b722d-327">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-328">ReadItem</span></span>|
|[<span data-ttu-id="b722d-329">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-330">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-330">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-331">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-331">Example</span></span>

<span data-ttu-id="b722d-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b722d-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="b722d-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b722d-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b722d-335">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="b722d-335">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b722d-336">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-336">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-337">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-337">Type:</span></span>

*   [<span data-ttu-id="b722d-338">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b722d-338">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b722d-339">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-339">Requirements</span></span>

|<span data-ttu-id="b722d-340">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-340">Requirement</span></span>| <span data-ttu-id="b722d-341">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-342">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-343">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-343">1.0</span></span>|
|[<span data-ttu-id="b722d-344">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-345">ReadItem</span></span>|
|[<span data-ttu-id="b722d-346">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-347">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-348">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-348">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="b722d-349">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="b722d-349">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="b722d-350">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-350">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-351">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-351">Read mode</span></span>

<span data-ttu-id="b722d-352">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-352">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-353">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-353">Compose mode</span></span>

<span data-ttu-id="b722d-354">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-354">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-355">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-355">Type:</span></span>

*   <span data-ttu-id="b722d-356">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="b722d-356">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-357">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-357">Requirements</span></span>

|<span data-ttu-id="b722d-358">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-358">Requirement</span></span>| <span data-ttu-id="b722d-359">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-360">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-361">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-361">1.0</span></span>|
|[<span data-ttu-id="b722d-362">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-363">ReadItem</span></span>|
|[<span data-ttu-id="b722d-364">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-365">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-366">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-366">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b722d-367">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b722d-367">normalizedSubject :String</span></span>

<span data-ttu-id="b722d-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b722d-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject).</span><span class="sxs-lookup"><span data-stu-id="b722d-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-372">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-372">Type:</span></span>

*   <span data-ttu-id="b722d-373">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-374">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-374">Requirements</span></span>

|<span data-ttu-id="b722d-375">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-375">Requirement</span></span>| <span data-ttu-id="b722d-376">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-377">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-378">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-378">1.0</span></span>|
|[<span data-ttu-id="b722d-379">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-380">ReadItem</span></span>|
|[<span data-ttu-id="b722d-381">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-382">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-383">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-383">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="b722d-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b722d-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="b722d-385">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-385">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-386">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-386">Type:</span></span>

*   [<span data-ttu-id="b722d-387">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b722d-387">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b722d-388">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-388">Requirements</span></span>

|<span data-ttu-id="b722d-389">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-389">Requirement</span></span>| <span data-ttu-id="b722d-390">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-390">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-391">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-392">1.3</span><span class="sxs-lookup"><span data-stu-id="b722d-392">1.3</span></span>|
|[<span data-ttu-id="b722d-393">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-393">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-394">ReadItem</span></span>|
|[<span data-ttu-id="b722d-395">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-395">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-396">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-396">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b722d-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b722d-398">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="b722d-398">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b722d-399">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-399">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-400">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-400">Read mode</span></span>

<span data-ttu-id="b722d-401">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="b722d-401">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-402">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-402">Compose mode</span></span>

<span data-ttu-id="b722d-403">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="b722d-403">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-404">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-404">Type:</span></span>

*   <span data-ttu-id="b722d-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-406">Requirements</span></span>

|<span data-ttu-id="b722d-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-407">Requirement</span></span>| <span data-ttu-id="b722d-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-410">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-410">1.0</span></span>|
|[<span data-ttu-id="b722d-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-412">ReadItem</span></span>|
|[<span data-ttu-id="b722d-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-414">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-414">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-415">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="b722d-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b722d-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="b722d-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-419">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-419">Type:</span></span>

*   [<span data-ttu-id="b722d-420">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b722d-420">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b722d-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-421">Requirements</span></span>

|<span data-ttu-id="b722d-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-422">Requirement</span></span>| <span data-ttu-id="b722d-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-425">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-425">1.0</span></span>|
|[<span data-ttu-id="b722d-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-427">ReadItem</span></span>|
|[<span data-ttu-id="b722d-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-429">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-430">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b722d-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b722d-432">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="b722d-432">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b722d-433">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-433">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-434">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-434">Read mode</span></span>

<span data-ttu-id="b722d-435">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="b722d-435">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-436">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-436">Compose mode</span></span>

<span data-ttu-id="b722d-437">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="b722d-437">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-438">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-438">Type:</span></span>

*   <span data-ttu-id="b722d-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-440">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-440">Requirements</span></span>

|<span data-ttu-id="b722d-441">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-441">Requirement</span></span>| <span data-ttu-id="b722d-442">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-443">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-444">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-444">1.0</span></span>|
|[<span data-ttu-id="b722d-445">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-446">ReadItem</span></span>|
|[<span data-ttu-id="b722d-447">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-448">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-449">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-449">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="b722d-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b722d-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="b722d-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b722d-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b722d-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="b722d-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-455">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b722d-455">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-456">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-456">Type:</span></span>

*   [<span data-ttu-id="b722d-457">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b722d-457">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b722d-458">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-458">Requirements</span></span>

|<span data-ttu-id="b722d-459">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-459">Requirement</span></span>| <span data-ttu-id="b722d-460">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-461">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-462">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-462">1.0</span></span>|
|[<span data-ttu-id="b722d-463">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-464">ReadItem</span></span>|
|[<span data-ttu-id="b722d-465">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-466">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-466">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-467">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-467">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="b722d-468">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b722d-468">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="b722d-469">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-469">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b722d-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="b722d-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-472">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-472">Read mode</span></span>

<span data-ttu-id="b722d-473">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="b722d-473">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-474">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-474">Compose mode</span></span>

<span data-ttu-id="b722d-475">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b722d-475">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b722d-476">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="b722d-476">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-477">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-477">Type:</span></span>

*   <span data-ttu-id="b722d-478">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b722d-478">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-479">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-479">Requirements</span></span>

|<span data-ttu-id="b722d-480">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-480">Requirement</span></span>| <span data-ttu-id="b722d-481">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-482">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-482">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-483">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-483">1.0</span></span>|
|[<span data-ttu-id="b722d-484">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-485">ReadItem</span></span>|
|[<span data-ttu-id="b722d-486">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-487">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-487">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-488">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-488">Example</span></span>

<span data-ttu-id="b722d-489">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="b722d-489">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="b722d-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b722d-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="b722d-491">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b722d-492">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="b722d-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-493">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-493">Read mode</span></span>

<span data-ttu-id="b722d-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="b722d-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b722d-496">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-496">Compose mode</span></span>

<span data-ttu-id="b722d-497">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="b722d-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b722d-498">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-498">Type:</span></span>

*   <span data-ttu-id="b722d-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b722d-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-500">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-500">Requirements</span></span>

|<span data-ttu-id="b722d-501">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-501">Requirement</span></span>| <span data-ttu-id="b722d-502">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-503">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-504">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-504">1.0</span></span>|
|[<span data-ttu-id="b722d-505">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-506">ReadItem</span></span>|
|[<span data-ttu-id="b722d-507">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-508">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-508">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b722d-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b722d-510">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="b722d-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b722d-511">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b722d-512">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-512">Read mode</span></span>

<span data-ttu-id="b722d-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="b722d-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b722d-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="b722d-515">Compose mode</span></span>

<span data-ttu-id="b722d-516">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="b722d-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b722d-517">Type :</span><span class="sxs-lookup"><span data-stu-id="b722d-517">Type:</span></span>

*   <span data-ttu-id="b722d-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b722d-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-519">Requirements</span></span>

|<span data-ttu-id="b722d-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-520">Requirement</span></span>| <span data-ttu-id="b722d-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-523">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-523">1.0</span></span>|
|[<span data-ttu-id="b722d-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-525">ReadItem</span></span>|
|[<span data-ttu-id="b722d-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-527">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-528">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-528">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b722d-529">Méthodes</span><span class="sxs-lookup"><span data-stu-id="b722d-529">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b722d-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b722d-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b722d-531">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="b722d-531">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b722d-532">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="b722d-532">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b722d-533">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="b722d-533">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-534">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-534">Parameters:</span></span>

|<span data-ttu-id="b722d-535">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-535">Name</span></span>| <span data-ttu-id="b722d-536">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-536">Type</span></span>| <span data-ttu-id="b722d-537">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-537">Attributes</span></span>| <span data-ttu-id="b722d-538">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-538">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b722d-539">String</span><span class="sxs-lookup"><span data-stu-id="b722d-539">String</span></span>||<span data-ttu-id="b722d-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b722d-542">String</span><span class="sxs-lookup"><span data-stu-id="b722d-542">String</span></span>||<span data-ttu-id="b722d-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b722d-545">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-545">Object</span></span>| <span data-ttu-id="b722d-546">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-546">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-547">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b722d-547">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b722d-548">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-548">Object</span></span>| <span data-ttu-id="b722d-549">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-549">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-550">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-550">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b722d-551">fonction</span><span class="sxs-lookup"><span data-stu-id="b722d-551">function</span></span>| <span data-ttu-id="b722d-552">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-552">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-553">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b722d-554">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b722d-554">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b722d-555">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b722d-555">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b722d-556">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b722d-556">Errors</span></span>

| <span data-ttu-id="b722d-557">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b722d-557">Error code</span></span> | <span data-ttu-id="b722d-558">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-558">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b722d-559">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="b722d-559">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b722d-560">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="b722d-560">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b722d-561">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b722d-561">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b722d-562">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-562">Requirements</span></span>

|<span data-ttu-id="b722d-563">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-563">Requirement</span></span>| <span data-ttu-id="b722d-564">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-565">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-566">1.1</span><span class="sxs-lookup"><span data-stu-id="b722d-566">1.1</span></span>|
|[<span data-ttu-id="b722d-567">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-568">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b722d-568">ReadWriteItem</span></span>|
|[<span data-ttu-id="b722d-569">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-570">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-570">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-571">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-571">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b722d-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b722d-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b722d-573">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-573">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b722d-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b722d-577">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="b722d-577">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b722d-578">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="b722d-578">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-579">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-579">Parameters:</span></span>

|<span data-ttu-id="b722d-580">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-580">Name</span></span>| <span data-ttu-id="b722d-581">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-581">Type</span></span>| <span data-ttu-id="b722d-582">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-582">Attributes</span></span>| <span data-ttu-id="b722d-583">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-583">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b722d-584">String</span><span class="sxs-lookup"><span data-stu-id="b722d-584">String</span></span>||<span data-ttu-id="b722d-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b722d-587">String</span><span class="sxs-lookup"><span data-stu-id="b722d-587">String</span></span>||<span data-ttu-id="b722d-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b722d-590">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-590">Object</span></span>| <span data-ttu-id="b722d-591">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-591">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-592">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b722d-592">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b722d-593">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-593">Object</span></span>| <span data-ttu-id="b722d-594">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-594">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-595">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-595">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b722d-596">fonction</span><span class="sxs-lookup"><span data-stu-id="b722d-596">function</span></span>| <span data-ttu-id="b722d-597">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-597">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-598">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b722d-599">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b722d-599">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b722d-600">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b722d-600">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b722d-601">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b722d-601">Errors</span></span>

| <span data-ttu-id="b722d-602">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b722d-602">Error code</span></span> | <span data-ttu-id="b722d-603">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-603">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b722d-604">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="b722d-604">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b722d-605">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-605">Requirements</span></span>

|<span data-ttu-id="b722d-606">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-606">Requirement</span></span>| <span data-ttu-id="b722d-607">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-609">1.1</span><span class="sxs-lookup"><span data-stu-id="b722d-609">1.1</span></span>|
|[<span data-ttu-id="b722d-610">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b722d-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="b722d-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-613">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-613">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-614">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-614">Example</span></span>

<span data-ttu-id="b722d-615">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="b722d-615">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="b722d-616">close()</span><span class="sxs-lookup"><span data-stu-id="b722d-616">close()</span></span>

<span data-ttu-id="b722d-617">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="b722d-617">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b722d-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="b722d-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-620">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-620">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b722d-621">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="b722d-621">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-622">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-622">Requirements</span></span>

|<span data-ttu-id="b722d-623">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-623">Requirement</span></span>| <span data-ttu-id="b722d-624">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-624">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-625">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-625">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-626">1.3</span><span class="sxs-lookup"><span data-stu-id="b722d-626">1.3</span></span>|
|[<span data-ttu-id="b722d-627">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-627">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-628">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b722d-628">Restricted</span></span>|
|[<span data-ttu-id="b722d-629">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-629">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-630">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-630">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b722d-631">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b722d-631">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b722d-632">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b722d-632">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-633">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-633">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b722d-634">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="b722d-634">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b722d-635">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="b722d-635">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b722d-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="b722d-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-639">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-639">Parameters:</span></span>

|<span data-ttu-id="b722d-640">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-640">Name</span></span>| <span data-ttu-id="b722d-641">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-641">Type</span></span>| <span data-ttu-id="b722d-642">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-642">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b722d-643">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b722d-643">String &#124; Object</span></span>| |<span data-ttu-id="b722d-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b722d-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b722d-646">**OU**</span><span class="sxs-lookup"><span data-stu-id="b722d-646">**OR**</span></span><br/><span data-ttu-id="b722d-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="b722d-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b722d-649">String</span><span class="sxs-lookup"><span data-stu-id="b722d-649">String</span></span> | <span data-ttu-id="b722d-650">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-650">&lt;optional&gt;</span></span> | <span data-ttu-id="b722d-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b722d-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b722d-653">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-653">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b722d-654">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-654">&lt;optional&gt;</span></span> | <span data-ttu-id="b722d-655">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-655">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b722d-656">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-656">String</span></span> | | <span data-ttu-id="b722d-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b722d-659">String</span><span class="sxs-lookup"><span data-stu-id="b722d-659">String</span></span> | | <span data-ttu-id="b722d-660">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-660">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b722d-661">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-661">String</span></span> | | <span data-ttu-id="b722d-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="b722d-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b722d-664">String</span><span class="sxs-lookup"><span data-stu-id="b722d-664">String</span></span> | | <span data-ttu-id="b722d-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b722d-668">function</span><span class="sxs-lookup"><span data-stu-id="b722d-668">function</span></span> | <span data-ttu-id="b722d-669">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-669">&lt;optional&gt;</span></span> | <span data-ttu-id="b722d-670">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b722d-671">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-671">Requirements</span></span>

|<span data-ttu-id="b722d-672">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-672">Requirement</span></span>| <span data-ttu-id="b722d-673">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-674">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-675">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-675">1.0</span></span>|
|[<span data-ttu-id="b722d-676">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-676">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-677">ReadItem</span></span>|
|[<span data-ttu-id="b722d-678">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-678">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-679">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-679">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b722d-680">Exemples</span><span class="sxs-lookup"><span data-stu-id="b722d-680">Examples</span></span>

<span data-ttu-id="b722d-681">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b722d-681">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b722d-682">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="b722d-682">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b722d-683">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="b722d-683">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b722d-684">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="b722d-684">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b722d-685">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-685">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b722d-686">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-686">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b722d-687">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b722d-687">displayReplyForm(formData)</span></span>

<span data-ttu-id="b722d-688">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b722d-688">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-689">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-689">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b722d-690">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="b722d-690">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b722d-691">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="b722d-691">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b722d-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="b722d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-695">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-695">Parameters:</span></span>

|<span data-ttu-id="b722d-696">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-696">Name</span></span>| <span data-ttu-id="b722d-697">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-697">Type</span></span>| <span data-ttu-id="b722d-698">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-698">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b722d-699">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b722d-699">String &#124; Object</span></span>| | <span data-ttu-id="b722d-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b722d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b722d-702">**OU**</span><span class="sxs-lookup"><span data-stu-id="b722d-702">**OR**</span></span><br/><span data-ttu-id="b722d-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="b722d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b722d-705">String</span><span class="sxs-lookup"><span data-stu-id="b722d-705">String</span></span> | <span data-ttu-id="b722d-706">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-706">&lt;optional&gt;</span></span> | <span data-ttu-id="b722d-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b722d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b722d-709">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-709">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b722d-710">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-710">&lt;optional&gt;</span></span> | <span data-ttu-id="b722d-711">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-711">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b722d-712">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-712">String</span></span> | | <span data-ttu-id="b722d-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b722d-715">String</span><span class="sxs-lookup"><span data-stu-id="b722d-715">String</span></span> | | <span data-ttu-id="b722d-716">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-716">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b722d-717">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-717">String</span></span> | | <span data-ttu-id="b722d-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="b722d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b722d-720">String</span><span class="sxs-lookup"><span data-stu-id="b722d-720">String</span></span> | | <span data-ttu-id="b722d-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="b722d-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b722d-724">function</span><span class="sxs-lookup"><span data-stu-id="b722d-724">function</span></span> | <span data-ttu-id="b722d-725">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-725">&lt;optional&gt;</span></span> | <span data-ttu-id="b722d-726">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-726">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b722d-727">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-727">Requirements</span></span>

|<span data-ttu-id="b722d-728">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-728">Requirement</span></span>| <span data-ttu-id="b722d-729">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-729">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-730">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-730">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-731">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-731">1.0</span></span>|
|[<span data-ttu-id="b722d-732">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-732">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-733">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-733">ReadItem</span></span>|
|[<span data-ttu-id="b722d-734">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-734">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-735">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-735">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b722d-736">Exemples</span><span class="sxs-lookup"><span data-stu-id="b722d-736">Examples</span></span>

<span data-ttu-id="b722d-737">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b722d-737">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b722d-738">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="b722d-738">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b722d-739">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="b722d-739">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b722d-740">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="b722d-740">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b722d-741">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-741">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b722d-742">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-742">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="b722d-743">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b722d-743">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="b722d-744">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b722d-744">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-745">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-745">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-746">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-746">Requirements</span></span>

|<span data-ttu-id="b722d-747">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-747">Requirement</span></span>| <span data-ttu-id="b722d-748">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-749">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-750">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-750">1.0</span></span>|
|[<span data-ttu-id="b722d-751">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-752">ReadItem</span></span>|
|[<span data-ttu-id="b722d-753">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-754">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b722d-755">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b722d-755">Returns:</span></span>

<span data-ttu-id="b722d-756">Type : [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b722d-756">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b722d-757">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-757">Example</span></span>

<span data-ttu-id="b722d-758">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-758">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="b722d-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b722d-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b722d-760">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b722d-760">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-761">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-762">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-762">Parameters:</span></span>

|<span data-ttu-id="b722d-763">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-763">Name</span></span>| <span data-ttu-id="b722d-764">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-764">Type</span></span>| <span data-ttu-id="b722d-765">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-765">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b722d-766">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b722d-766">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="b722d-767">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="b722d-767">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b722d-768">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-768">Requirements</span></span>

|<span data-ttu-id="b722d-769">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-769">Requirement</span></span>| <span data-ttu-id="b722d-770">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-771">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-772">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-772">1.0</span></span>|
|[<span data-ttu-id="b722d-773">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-774">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b722d-774">Restricted</span></span>|
|[<span data-ttu-id="b722d-775">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-776">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b722d-777">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b722d-777">Returns:</span></span>

<span data-ttu-id="b722d-778">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="b722d-778">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b722d-779">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="b722d-779">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b722d-780">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="b722d-780">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b722d-781">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="b722d-781">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b722d-782">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="b722d-782">Value of `entityType`</span></span> | <span data-ttu-id="b722d-783">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="b722d-783">Type of objects in returned array</span></span> | <span data-ttu-id="b722d-784">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="b722d-784">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b722d-785">String</span><span class="sxs-lookup"><span data-stu-id="b722d-785">String</span></span> | <span data-ttu-id="b722d-786">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b722d-786">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b722d-787">Contact</span><span class="sxs-lookup"><span data-stu-id="b722d-787">Contact</span></span> | <span data-ttu-id="b722d-788">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b722d-788">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b722d-789">String</span><span class="sxs-lookup"><span data-stu-id="b722d-789">String</span></span> | <span data-ttu-id="b722d-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b722d-790">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b722d-791">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b722d-791">MeetingSuggestion</span></span> | <span data-ttu-id="b722d-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b722d-792">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b722d-793">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b722d-793">PhoneNumber</span></span> | <span data-ttu-id="b722d-794">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b722d-794">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b722d-795">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b722d-795">TaskSuggestion</span></span> | <span data-ttu-id="b722d-796">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b722d-796">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b722d-797">String</span><span class="sxs-lookup"><span data-stu-id="b722d-797">String</span></span> | <span data-ttu-id="b722d-798">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b722d-798">**Restricted**</span></span> |

<span data-ttu-id="b722d-799">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b722d-799">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b722d-800">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-800">Example</span></span>

<span data-ttu-id="b722d-801">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="b722d-801">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="b722d-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b722d-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b722d-803">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b722d-803">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-804">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-804">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b722d-805">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="b722d-805">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-806">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-806">Parameters:</span></span>

|<span data-ttu-id="b722d-807">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-807">Name</span></span>| <span data-ttu-id="b722d-808">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-808">Type</span></span>| <span data-ttu-id="b722d-809">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-809">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b722d-810">String</span><span class="sxs-lookup"><span data-stu-id="b722d-810">String</span></span>|<span data-ttu-id="b722d-811">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="b722d-811">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b722d-812">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-812">Requirements</span></span>

|<span data-ttu-id="b722d-813">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-813">Requirement</span></span>| <span data-ttu-id="b722d-814">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-814">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-815">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-815">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-816">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-816">1.0</span></span>|
|[<span data-ttu-id="b722d-817">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-817">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-818">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-818">ReadItem</span></span>|
|[<span data-ttu-id="b722d-819">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-819">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-820">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-820">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b722d-821">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b722d-821">Returns:</span></span>

<span data-ttu-id="b722d-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="b722d-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b722d-824">Type : Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b722d-824">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b722d-825">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b722d-825">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b722d-826">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b722d-826">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-827">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-827">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b722d-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="b722d-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b722d-831">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="b722d-831">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b722d-832">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b722d-832">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b722d-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b722d-836">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-836">Requirements</span></span>

|<span data-ttu-id="b722d-837">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-837">Requirement</span></span>| <span data-ttu-id="b722d-838">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-839">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-840">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-840">1.0</span></span>|
|[<span data-ttu-id="b722d-841">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-842">ReadItem</span></span>|
|[<span data-ttu-id="b722d-843">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-844">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b722d-845">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b722d-845">Returns:</span></span>

<span data-ttu-id="b722d-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="b722d-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b722d-848">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b722d-848">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b722d-849">Object</span><span class="sxs-lookup"><span data-stu-id="b722d-849">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b722d-850">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-850">Example</span></span>

<span data-ttu-id="b722d-851">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="b722d-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b722d-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="b722d-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b722d-853">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b722d-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-854">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b722d-854">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b722d-855">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="b722d-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b722d-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="b722d-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-858">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-858">Parameters:</span></span>

|<span data-ttu-id="b722d-859">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-859">Name</span></span>| <span data-ttu-id="b722d-860">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-860">Type</span></span>| <span data-ttu-id="b722d-861">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b722d-862">String</span><span class="sxs-lookup"><span data-stu-id="b722d-862">String</span></span>|<span data-ttu-id="b722d-863">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="b722d-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b722d-864">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-864">Requirements</span></span>

|<span data-ttu-id="b722d-865">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-865">Requirement</span></span>| <span data-ttu-id="b722d-866">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-867">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-868">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-868">1.0</span></span>|
|[<span data-ttu-id="b722d-869">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-869">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-870">ReadItem</span></span>|
|[<span data-ttu-id="b722d-871">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-871">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-872">Lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b722d-873">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b722d-873">Returns:</span></span>

<span data-ttu-id="b722d-874">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="b722d-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b722d-875">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b722d-875">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b722d-876">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b722d-876">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b722d-877">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-877">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b722d-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b722d-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b722d-879">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="b722d-879">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b722d-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="b722d-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-882">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-882">Parameters:</span></span>

|<span data-ttu-id="b722d-883">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-883">Name</span></span>| <span data-ttu-id="b722d-884">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-884">Type</span></span>| <span data-ttu-id="b722d-885">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-885">Attributes</span></span>| <span data-ttu-id="b722d-886">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-886">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b722d-887">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b722d-887">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b722d-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="b722d-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b722d-891">Object</span><span class="sxs-lookup"><span data-stu-id="b722d-891">Object</span></span>| <span data-ttu-id="b722d-892">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-892">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-893">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b722d-893">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b722d-894">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-894">Object</span></span>| <span data-ttu-id="b722d-895">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-895">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-896">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-896">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b722d-897">fonction</span><span class="sxs-lookup"><span data-stu-id="b722d-897">function</span></span>||<span data-ttu-id="b722d-898">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-898">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b722d-899">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="b722d-899">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b722d-900">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="b722d-900">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b722d-901">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-901">Requirements</span></span>

|<span data-ttu-id="b722d-902">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-902">Requirement</span></span>| <span data-ttu-id="b722d-903">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-904">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-905">1.2</span><span class="sxs-lookup"><span data-stu-id="b722d-905">1.2</span></span>|
|[<span data-ttu-id="b722d-906">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-907">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b722d-907">ReadWriteItem</span></span>|
|[<span data-ttu-id="b722d-908">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-909">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-909">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b722d-910">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b722d-910">Returns:</span></span>

<span data-ttu-id="b722d-911">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="b722d-911">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b722d-912">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="b722d-912">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b722d-913">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b722d-913">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b722d-914">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-914">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b722d-915">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b722d-915">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b722d-916">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="b722d-916">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b722d-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="b722d-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-920">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-920">Parameters:</span></span>

|<span data-ttu-id="b722d-921">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-921">Name</span></span>| <span data-ttu-id="b722d-922">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-922">Type</span></span>| <span data-ttu-id="b722d-923">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-923">Attributes</span></span>| <span data-ttu-id="b722d-924">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-924">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b722d-925">function</span><span class="sxs-lookup"><span data-stu-id="b722d-925">function</span></span>||<span data-ttu-id="b722d-926">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-926">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b722d-927">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b722d-927">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b722d-928">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="b722d-928">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b722d-929">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-929">Object</span></span>| <span data-ttu-id="b722d-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-930">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-931">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-931">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b722d-932">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-932">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b722d-933">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-933">Requirements</span></span>

|<span data-ttu-id="b722d-934">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-934">Requirement</span></span>| <span data-ttu-id="b722d-935">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-936">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-937">1.0</span><span class="sxs-lookup"><span data-stu-id="b722d-937">1.0</span></span>|
|[<span data-ttu-id="b722d-938">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b722d-939">ReadItem</span></span>|
|[<span data-ttu-id="b722d-940">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-941">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b722d-941">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-942">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-942">Example</span></span>

<span data-ttu-id="b722d-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b722d-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b722d-946">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b722d-946">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b722d-947">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b722d-947">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b722d-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="b722d-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-952">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-952">Parameters:</span></span>

|<span data-ttu-id="b722d-953">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-953">Name</span></span>| <span data-ttu-id="b722d-954">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-954">Type</span></span>| <span data-ttu-id="b722d-955">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-955">Attributes</span></span>| <span data-ttu-id="b722d-956">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-956">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b722d-957">String</span><span class="sxs-lookup"><span data-stu-id="b722d-957">String</span></span>||<span data-ttu-id="b722d-958">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="b722d-958">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="b722d-959">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-959">Object</span></span>| <span data-ttu-id="b722d-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-960">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-961">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b722d-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b722d-962">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-962">Object</span></span>| <span data-ttu-id="b722d-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-963">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-964">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b722d-965">fonction</span><span class="sxs-lookup"><span data-stu-id="b722d-965">function</span></span>| <span data-ttu-id="b722d-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-966">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-967">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b722d-968">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="b722d-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b722d-969">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b722d-969">Errors</span></span>

| <span data-ttu-id="b722d-970">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b722d-970">Error code</span></span> | <span data-ttu-id="b722d-971">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b722d-972">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="b722d-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b722d-973">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-973">Requirements</span></span>

|<span data-ttu-id="b722d-974">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-974">Requirement</span></span>| <span data-ttu-id="b722d-975">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-976">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-977">1.1</span><span class="sxs-lookup"><span data-stu-id="b722d-977">1.1</span></span>|
|[<span data-ttu-id="b722d-978">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b722d-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="b722d-980">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-981">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-982">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-982">Example</span></span>

<span data-ttu-id="b722d-983">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="b722d-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b722d-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b722d-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="b722d-985">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b722d-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="b722d-p166">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="b722d-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-989">si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="b722d-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="b722d-990">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="b722d-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b722d-p168">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="b722d-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b722d-994">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="b722d-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b722d-995">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="b722d-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="b722d-996">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="b722d-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b722d-997">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="b722d-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-998">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-998">Parameters:</span></span>

|<span data-ttu-id="b722d-999">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-999">Name</span></span>| <span data-ttu-id="b722d-1000">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-1000">Type</span></span>| <span data-ttu-id="b722d-1001">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-1001">Attributes</span></span>| <span data-ttu-id="b722d-1002">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="b722d-1003">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-1003">Object</span></span>| <span data-ttu-id="b722d-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-1005">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b722d-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b722d-1006">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-1006">Object</span></span>| <span data-ttu-id="b722d-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-1008">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b722d-1009">fonction</span><span class="sxs-lookup"><span data-stu-id="b722d-1009">function</span></span>||<span data-ttu-id="b722d-1010">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b722d-1011">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b722d-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b722d-1012">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-1012">Requirements</span></span>

|<span data-ttu-id="b722d-1013">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-1013">Requirement</span></span>| <span data-ttu-id="b722d-1014">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-1015">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="b722d-1016">1.3</span></span>|
|[<span data-ttu-id="b722d-1017">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b722d-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="b722d-1019">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-1020">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b722d-1021">範例</span><span class="sxs-lookup"><span data-stu-id="b722d-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="b722d-p170">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b722d-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b722d-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b722d-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b722d-1025">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b722d-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b722d-p171">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="b722d-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b722d-1029">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="b722d-1029">Parameters:</span></span>

|<span data-ttu-id="b722d-1030">Nom</span><span class="sxs-lookup"><span data-stu-id="b722d-1030">Name</span></span>| <span data-ttu-id="b722d-1031">Type</span><span class="sxs-lookup"><span data-stu-id="b722d-1031">Type</span></span>| <span data-ttu-id="b722d-1032">Attributs</span><span class="sxs-lookup"><span data-stu-id="b722d-1032">Attributes</span></span>| <span data-ttu-id="b722d-1033">Description</span><span class="sxs-lookup"><span data-stu-id="b722d-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b722d-1034">String</span><span class="sxs-lookup"><span data-stu-id="b722d-1034">String</span></span>||<span data-ttu-id="b722d-p172">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="b722d-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b722d-1038">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-1038">Object</span></span>| <span data-ttu-id="b722d-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-1040">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="b722d-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b722d-1041">Objet</span><span class="sxs-lookup"><span data-stu-id="b722d-1041">Object</span></span>| <span data-ttu-id="b722d-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-1043">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="b722d-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b722d-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b722d-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b722d-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b722d-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="b722d-p173">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="b722d-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b722d-p174">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="b722d-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b722d-1050">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="b722d-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b722d-1051">fonction</span><span class="sxs-lookup"><span data-stu-id="b722d-1051">function</span></span>||<span data-ttu-id="b722d-1052">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b722d-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b722d-1053">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b722d-1053">Requirements</span></span>

|<span data-ttu-id="b722d-1054">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b722d-1054">Requirement</span></span>| <span data-ttu-id="b722d-1055">Valeur</span><span class="sxs-lookup"><span data-stu-id="b722d-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="b722d-1056">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b722d-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b722d-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="b722d-1057">1.2</span></span>|
|[<span data-ttu-id="b722d-1058">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b722d-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b722d-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b722d-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="b722d-1060">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b722d-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b722d-1061">Composition</span><span class="sxs-lookup"><span data-stu-id="b722d-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b722d-1062">Exemple</span><span class="sxs-lookup"><span data-stu-id="b722d-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
