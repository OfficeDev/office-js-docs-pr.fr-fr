---
title: Office.Context.Mailbox.Item - exigence défini 1.1
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: 63460494a049bb83d3af69f6808396e426842f1e
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389577"
---
# <a name="item"></a><span data-ttu-id="9a698-102">élément</span><span class="sxs-lookup"><span data-stu-id="9a698-102">item</span></span>

### <span data-ttu-id="9a698-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="9a698-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="9a698-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9a698-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-107">Requirements</span></span>

|<span data-ttu-id="9a698-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-108">Requirement</span></span>| <span data-ttu-id="9a698-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-111">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-111">1.0</span></span>|
|[<span data-ttu-id="9a698-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="9a698-113">Restricted</span></span>|
|[<span data-ttu-id="9a698-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-115">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="9a698-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-116">Example</span></span>

<span data-ttu-id="9a698-117">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a698-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9a698-118">Membres</span><span class="sxs-lookup"><span data-stu-id="9a698-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="9a698-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9a698-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="9a698-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-122">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="9a698-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9a698-123">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9a698-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-124">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-124">Type:</span></span>

*   <span data-ttu-id="9a698-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9a698-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-126">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-126">Requirements</span></span>

|<span data-ttu-id="9a698-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-127">Requirement</span></span>| <span data-ttu-id="9a698-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-129">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-130">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-130">1.0</span></span>|
|[<span data-ttu-id="9a698-131">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-132">ReadItem</span></span>|
|[<span data-ttu-id="9a698-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-134">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-135">Example</span></span>

<span data-ttu-id="9a698-136">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9a698-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9a698-138">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="9a698-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9a698-139">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-140">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-140">Type:</span></span>

*   [<span data-ttu-id="9a698-141">Destinataires</span><span class="sxs-lookup"><span data-stu-id="9a698-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9a698-142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-142">Requirements</span></span>

|<span data-ttu-id="9a698-143">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-143">Requirement</span></span>| <span data-ttu-id="9a698-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-145">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-146">1.1</span><span class="sxs-lookup"><span data-stu-id="9a698-146">1.1</span></span>|
|[<span data-ttu-id="9a698-147">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-148">ReadItem</span></span>|
|[<span data-ttu-id="9a698-149">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-150">Composition</span><span class="sxs-lookup"><span data-stu-id="9a698-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-151">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="9a698-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="9a698-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="9a698-153">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="9a698-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-154">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-154">Type:</span></span>

*   [<span data-ttu-id="9a698-155">Corps</span><span class="sxs-lookup"><span data-stu-id="9a698-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="9a698-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-156">Requirements</span></span>

|<span data-ttu-id="9a698-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-157">Requirement</span></span>| <span data-ttu-id="9a698-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-160">1.1</span><span class="sxs-lookup"><span data-stu-id="9a698-160">1.1</span></span>|
|[<span data-ttu-id="9a698-161">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-162">ReadItem</span></span>|
|[<span data-ttu-id="9a698-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9a698-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9a698-166">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="9a698-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9a698-167">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-168">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-168">Read mode</span></span>

<span data-ttu-id="9a698-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="9a698-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-171">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-171">Compose mode</span></span>

<span data-ttu-id="9a698-172">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="9a698-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-173">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-173">Type:</span></span>

*   <span data-ttu-id="9a698-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-175">Requirements</span></span>

|<span data-ttu-id="9a698-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-176">Requirement</span></span>| <span data-ttu-id="9a698-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-179">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-179">1.0</span></span>|
|[<span data-ttu-id="9a698-180">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-181">ReadItem</span></span>|
|[<span data-ttu-id="9a698-182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-183">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-184">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="9a698-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="9a698-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="9a698-186">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="9a698-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9a698-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="9a698-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9a698-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="9a698-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-191">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-191">Type:</span></span>

*   <span data-ttu-id="9a698-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a698-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-193">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-193">Requirements</span></span>

|<span data-ttu-id="9a698-194">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-194">Requirement</span></span>| <span data-ttu-id="9a698-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-196">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-197">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-197">1.0</span></span>|
|[<span data-ttu-id="9a698-198">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-199">ReadItem</span></span>|
|[<span data-ttu-id="9a698-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-201">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="9a698-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="9a698-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="9a698-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-205">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-205">Type:</span></span>

*   <span data-ttu-id="9a698-206">Date</span><span class="sxs-lookup"><span data-stu-id="9a698-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-207">Requirements</span></span>

|<span data-ttu-id="9a698-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-208">Requirement</span></span>| <span data-ttu-id="9a698-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-211">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-211">1.0</span></span>|
|[<span data-ttu-id="9a698-212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-213">ReadItem</span></span>|
|[<span data-ttu-id="9a698-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-215">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9a698-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="9a698-217">dateTimeModified :Date</span></span>

<span data-ttu-id="9a698-p111">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-220">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-221">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-221">Type:</span></span>

*   <span data-ttu-id="9a698-222">Date</span><span class="sxs-lookup"><span data-stu-id="9a698-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-223">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-223">Requirements</span></span>

|<span data-ttu-id="9a698-224">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-224">Requirement</span></span>| <span data-ttu-id="9a698-225">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-226">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-227">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-227">1.0</span></span>|
|[<span data-ttu-id="9a698-228">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-229">ReadItem</span></span>|
|[<span data-ttu-id="9a698-230">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-231">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-232">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="9a698-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a698-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="9a698-234">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9a698-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="9a698-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-237">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-237">Read mode</span></span>

<span data-ttu-id="9a698-238">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="9a698-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-239">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-239">Compose mode</span></span>

<span data-ttu-id="9a698-240">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a698-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9a698-241">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="9a698-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-242">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-242">Type:</span></span>

*   <span data-ttu-id="9a698-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a698-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-244">Requirements</span></span>

|<span data-ttu-id="9a698-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-245">Requirement</span></span>| <span data-ttu-id="9a698-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-248">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-248">1.0</span></span>|
|[<span data-ttu-id="9a698-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-250">ReadItem</span></span>|
|[<span data-ttu-id="9a698-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-252">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-253">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-253">Example</span></span>

<span data-ttu-id="9a698-254">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a698-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9a698-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9a698-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9a698-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9a698-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="9a698-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-260">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9a698-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-261">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-261">Type:</span></span>

*   [<span data-ttu-id="9a698-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a698-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9a698-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-263">Requirements</span></span>

|<span data-ttu-id="9a698-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-264">Requirement</span></span>| <span data-ttu-id="9a698-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-267">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-267">1.0</span></span>|
|[<span data-ttu-id="9a698-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-269">ReadItem</span></span>|
|[<span data-ttu-id="9a698-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-271">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="9a698-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="9a698-272">internetMessageId :String</span></span>

<span data-ttu-id="9a698-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-275">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-275">Type:</span></span>

*   <span data-ttu-id="9a698-276">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a698-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-277">Requirements</span></span>

|<span data-ttu-id="9a698-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-278">Requirement</span></span>| <span data-ttu-id="9a698-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-281">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-281">1.0</span></span>|
|[<span data-ttu-id="9a698-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-283">ReadItem</span></span>|
|[<span data-ttu-id="9a698-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-285">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9a698-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="9a698-287">itemClass :String</span></span>

<span data-ttu-id="9a698-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9a698-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9a698-292">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-292">Type</span></span> | <span data-ttu-id="9a698-293">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-293">Description</span></span> | <span data-ttu-id="9a698-294">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="9a698-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9a698-295">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="9a698-295">Appointment items</span></span> | <span data-ttu-id="9a698-296">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="9a698-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="9a698-297">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="9a698-297">Message items</span></span> | <span data-ttu-id="9a698-298">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="9a698-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9a698-299">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="9a698-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-300">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-300">Type:</span></span>

*   <span data-ttu-id="9a698-301">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a698-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-302">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-302">Requirements</span></span>

|<span data-ttu-id="9a698-303">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-303">Requirement</span></span>| <span data-ttu-id="9a698-304">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-305">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-306">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-306">1.0</span></span>|
|[<span data-ttu-id="9a698-307">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-308">ReadItem</span></span>|
|[<span data-ttu-id="9a698-309">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-310">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-311">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9a698-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="9a698-312">(nullable) itemId :String</span></span>

<span data-ttu-id="9a698-p118">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-315">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="9a698-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9a698-316">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a698-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9a698-317">Avant d’effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande `Office.context.mailbox.convertToRestId`, qui est disponible à partir de l’ensemble de conditions requises 1.3.</span><span class="sxs-lookup"><span data-stu-id="9a698-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="9a698-318">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9a698-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-319">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-319">Type:</span></span>

*   <span data-ttu-id="9a698-320">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a698-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-321">Requirements</span></span>

|<span data-ttu-id="9a698-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-322">Requirement</span></span>| <span data-ttu-id="9a698-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-325">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-325">1.0</span></span>|
|[<span data-ttu-id="9a698-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-327">ReadItem</span></span>|
|[<span data-ttu-id="9a698-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-329">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-330">Example</span></span>

<span data-ttu-id="9a698-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="9a698-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="9a698-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9a698-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9a698-334">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="9a698-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9a698-335">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-336">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-336">Type:</span></span>

*   [<span data-ttu-id="9a698-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9a698-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9a698-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-338">Requirements</span></span>

|<span data-ttu-id="9a698-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-339">Requirement</span></span>| <span data-ttu-id="9a698-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-342">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-342">1.0</span></span>|
|[<span data-ttu-id="9a698-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-344">ReadItem</span></span>|
|[<span data-ttu-id="9a698-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-346">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="9a698-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="9a698-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="9a698-349">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-350">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-350">Read mode</span></span>

<span data-ttu-id="9a698-351">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-352">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-352">Compose mode</span></span>

<span data-ttu-id="9a698-353">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-354">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-354">Type:</span></span>

*   <span data-ttu-id="9a698-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="9a698-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-356">Requirements</span></span>

|<span data-ttu-id="9a698-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-357">Requirement</span></span>| <span data-ttu-id="9a698-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-359">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-360">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-360">1.0</span></span>|
|[<span data-ttu-id="9a698-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-362">ReadItem</span></span>|
|[<span data-ttu-id="9a698-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-364">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9a698-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="9a698-366">normalizedSubject :String</span></span>

<span data-ttu-id="9a698-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9a698-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="9a698-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-371">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-371">Type:</span></span>

*   <span data-ttu-id="9a698-372">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a698-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-373">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-373">Requirements</span></span>

|<span data-ttu-id="9a698-374">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-374">Requirement</span></span>| <span data-ttu-id="9a698-375">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-376">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-377">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-377">1.0</span></span>|
|[<span data-ttu-id="9a698-378">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-379">ReadItem</span></span>|
|[<span data-ttu-id="9a698-380">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-381">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-382">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9a698-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9a698-384">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="9a698-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9a698-385">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-386">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-386">Read mode</span></span>

<span data-ttu-id="9a698-387">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="9a698-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-388">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-388">Compose mode</span></span>

<span data-ttu-id="9a698-389">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="9a698-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-390">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-390">Type:</span></span>

*   <span data-ttu-id="9a698-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-392">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-392">Requirements</span></span>

|<span data-ttu-id="9a698-393">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-393">Requirement</span></span>| <span data-ttu-id="9a698-394">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-395">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-396">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-396">1.0</span></span>|
|[<span data-ttu-id="9a698-397">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-398">ReadItem</span></span>|
|[<span data-ttu-id="9a698-399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-400">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-401">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9a698-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9a698-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9a698-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-405">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-405">Type:</span></span>

*   [<span data-ttu-id="9a698-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a698-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9a698-407">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-407">Requirements</span></span>

|<span data-ttu-id="9a698-408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-408">Requirement</span></span>| <span data-ttu-id="9a698-409">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-410">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-411">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-411">1.0</span></span>|
|[<span data-ttu-id="9a698-412">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-413">ReadItem</span></span>|
|[<span data-ttu-id="9a698-414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-415">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-416">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9a698-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9a698-418">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="9a698-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9a698-419">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-420">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-420">Read mode</span></span>

<span data-ttu-id="9a698-421">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="9a698-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-422">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-422">Compose mode</span></span>

<span data-ttu-id="9a698-423">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="9a698-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-424">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-424">Type:</span></span>

*   <span data-ttu-id="9a698-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-426">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-426">Requirements</span></span>

|<span data-ttu-id="9a698-427">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-427">Requirement</span></span>| <span data-ttu-id="9a698-428">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-429">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-430">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-430">1.0</span></span>|
|[<span data-ttu-id="9a698-431">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-432">ReadItem</span></span>|
|[<span data-ttu-id="9a698-433">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-434">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-435">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="9a698-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9a698-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="9a698-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="9a698-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9a698-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="9a698-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-441">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9a698-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-442">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-442">Type:</span></span>

*   [<span data-ttu-id="9a698-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a698-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9a698-444">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-444">Requirements</span></span>

|<span data-ttu-id="9a698-445">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-445">Requirement</span></span>| <span data-ttu-id="9a698-446">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-447">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-448">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-448">1.0</span></span>|
|[<span data-ttu-id="9a698-449">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-450">ReadItem</span></span>|
|[<span data-ttu-id="9a698-451">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-452">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-453">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="9a698-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a698-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="9a698-455">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9a698-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="9a698-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-458">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-458">Read mode</span></span>

<span data-ttu-id="9a698-459">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="9a698-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-460">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-460">Compose mode</span></span>

<span data-ttu-id="9a698-461">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a698-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9a698-462">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="9a698-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-463">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-463">Type:</span></span>

*   <span data-ttu-id="9a698-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a698-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-465">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-465">Requirements</span></span>

|<span data-ttu-id="9a698-466">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-466">Requirement</span></span>| <span data-ttu-id="9a698-467">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-468">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-469">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-469">1.0</span></span>|
|[<span data-ttu-id="9a698-470">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-471">ReadItem</span></span>|
|[<span data-ttu-id="9a698-472">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-473">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-474">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-474">Example</span></span>

<span data-ttu-id="9a698-475">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="9a698-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="9a698-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9a698-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="9a698-477">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="9a698-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9a698-478">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="9a698-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-479">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-479">Read mode</span></span>

<span data-ttu-id="9a698-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9a698-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="9a698-482">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-482">Compose mode</span></span>

<span data-ttu-id="9a698-483">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="9a698-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a698-484">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-484">Type:</span></span>

*   <span data-ttu-id="9a698-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9a698-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-486">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-486">Requirements</span></span>

|<span data-ttu-id="9a698-487">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-487">Requirement</span></span>| <span data-ttu-id="9a698-488">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-489">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-490">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-490">1.0</span></span>|
|[<span data-ttu-id="9a698-491">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-492">ReadItem</span></span>|
|[<span data-ttu-id="9a698-493">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-494">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="9a698-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="9a698-496">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="9a698-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9a698-497">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a698-498">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-498">Read mode</span></span>

<span data-ttu-id="9a698-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="9a698-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a698-501">Mode composition</span><span class="sxs-lookup"><span data-stu-id="9a698-501">Compose mode</span></span>

<span data-ttu-id="9a698-502">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="9a698-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9a698-503">Type :</span><span class="sxs-lookup"><span data-stu-id="9a698-503">Type:</span></span>

*   <span data-ttu-id="9a698-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a698-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-505">Requirements</span></span>

|<span data-ttu-id="9a698-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-506">Requirement</span></span>| <span data-ttu-id="9a698-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-509">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-509">1.0</span></span>|
|[<span data-ttu-id="9a698-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-511">ReadItem</span></span>|
|[<span data-ttu-id="9a698-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-513">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="9a698-515">Méthodes</span><span class="sxs-lookup"><span data-stu-id="9a698-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9a698-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a698-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a698-517">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="9a698-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9a698-518">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="9a698-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9a698-519">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="9a698-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-520">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-520">Parameters:</span></span>

|<span data-ttu-id="9a698-521">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-521">Name</span></span>| <span data-ttu-id="9a698-522">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-522">Type</span></span>| <span data-ttu-id="9a698-523">Attributs</span><span class="sxs-lookup"><span data-stu-id="9a698-523">Attributes</span></span>| <span data-ttu-id="9a698-524">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9a698-525">String</span><span class="sxs-lookup"><span data-stu-id="9a698-525">String</span></span>||<span data-ttu-id="9a698-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="9a698-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9a698-528">String</span><span class="sxs-lookup"><span data-stu-id="9a698-528">String</span></span>||<span data-ttu-id="9a698-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="9a698-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9a698-531">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-531">Object</span></span>| <span data-ttu-id="9a698-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-532">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-533">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="9a698-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9a698-534">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-534">Object</span></span>| <span data-ttu-id="9a698-535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-535">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-536">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9a698-537">fonction</span><span class="sxs-lookup"><span data-stu-id="9a698-537">function</span></span>| <span data-ttu-id="9a698-538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-538">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-539">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a698-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a698-540">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a698-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a698-541">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="9a698-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a698-542">Erreurs</span><span class="sxs-lookup"><span data-stu-id="9a698-542">Errors</span></span>

| <span data-ttu-id="9a698-543">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="9a698-543">Error code</span></span> | <span data-ttu-id="9a698-544">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9a698-545">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="9a698-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9a698-546">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="9a698-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9a698-547">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="9a698-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9a698-548">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-548">Requirements</span></span>

|<span data-ttu-id="9a698-549">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-549">Requirement</span></span>| <span data-ttu-id="9a698-550">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-551">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-552">1.1</span><span class="sxs-lookup"><span data-stu-id="9a698-552">1.1</span></span>|
|[<span data-ttu-id="9a698-553">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a698-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a698-555">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-556">Composition</span><span class="sxs-lookup"><span data-stu-id="9a698-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-557">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9a698-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a698-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a698-559">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9a698-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9a698-563">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="9a698-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9a698-564">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="9a698-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-565">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-565">Parameters:</span></span>

|<span data-ttu-id="9a698-566">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-566">Name</span></span>| <span data-ttu-id="9a698-567">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-567">Type</span></span>| <span data-ttu-id="9a698-568">Attributs</span><span class="sxs-lookup"><span data-stu-id="9a698-568">Attributes</span></span>| <span data-ttu-id="9a698-569">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9a698-570">String</span><span class="sxs-lookup"><span data-stu-id="9a698-570">String</span></span>||<span data-ttu-id="9a698-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="9a698-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9a698-573">String</span><span class="sxs-lookup"><span data-stu-id="9a698-573">String</span></span>||<span data-ttu-id="9a698-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="9a698-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9a698-576">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-576">Object</span></span>| <span data-ttu-id="9a698-577">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-577">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-578">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="9a698-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9a698-579">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-579">Object</span></span>| <span data-ttu-id="9a698-580">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-580">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-581">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9a698-582">fonction</span><span class="sxs-lookup"><span data-stu-id="9a698-582">function</span></span>| <span data-ttu-id="9a698-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-583">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-584">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a698-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a698-585">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a698-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a698-586">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="9a698-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a698-587">Erreurs</span><span class="sxs-lookup"><span data-stu-id="9a698-587">Errors</span></span>

| <span data-ttu-id="9a698-588">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="9a698-588">Error code</span></span> | <span data-ttu-id="9a698-589">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9a698-590">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="9a698-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9a698-591">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-591">Requirements</span></span>

|<span data-ttu-id="9a698-592">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-592">Requirement</span></span>| <span data-ttu-id="9a698-593">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-594">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-595">1.1</span><span class="sxs-lookup"><span data-stu-id="9a698-595">1.1</span></span>|
|[<span data-ttu-id="9a698-596">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a698-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a698-598">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-599">Composition</span><span class="sxs-lookup"><span data-stu-id="9a698-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-600">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-600">Example</span></span>

<span data-ttu-id="9a698-601">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9a698-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="9a698-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9a698-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="9a698-603">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9a698-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-604">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9a698-605">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="9a698-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9a698-606">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="9a698-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-607">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyAllForm` n’est pas prise en charge dans l’ensemble des conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="9a698-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="9a698-608">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="9a698-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-609">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-609">Parameters:</span></span>

|<span data-ttu-id="9a698-610">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-610">Name</span></span>| <span data-ttu-id="9a698-611">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-611">Type</span></span>| <span data-ttu-id="9a698-612">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9a698-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9a698-613">String &#124; Object</span></span>| |<span data-ttu-id="9a698-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9a698-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9a698-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="9a698-616">**OR**</span></span><br/><span data-ttu-id="9a698-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="9a698-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9a698-619">String</span><span class="sxs-lookup"><span data-stu-id="9a698-619">String</span></span> | <span data-ttu-id="9a698-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-620">&lt;optional&gt;</span></span> | <span data-ttu-id="9a698-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9a698-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="9a698-623">function</span><span class="sxs-lookup"><span data-stu-id="9a698-623">function</span></span> | <span data-ttu-id="9a698-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-624">&lt;optional&gt;</span></span> | <span data-ttu-id="9a698-625">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a698-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9a698-626">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-626">Requirements</span></span>

|<span data-ttu-id="9a698-627">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-627">Requirement</span></span>| <span data-ttu-id="9a698-628">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-629">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-630">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-630">1.0</span></span>|
|[<span data-ttu-id="9a698-631">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-632">ReadItem</span></span>|
|[<span data-ttu-id="9a698-633">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-634">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a698-635">Exemples</span><span class="sxs-lookup"><span data-stu-id="9a698-635">Examples</span></span>

<span data-ttu-id="9a698-636">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9a698-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9a698-637">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="9a698-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9a698-638">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="9a698-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9a698-639">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-639">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="9a698-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9a698-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="9a698-641">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9a698-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-642">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9a698-643">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="9a698-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9a698-644">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="9a698-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-645">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyForm` n’est pas prise en charge dans l’ensemble des conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="9a698-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="9a698-646">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="9a698-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-647">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-647">Parameters:</span></span>

|<span data-ttu-id="9a698-648">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-648">Name</span></span>| <span data-ttu-id="9a698-649">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-649">Type</span></span>| <span data-ttu-id="9a698-650">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9a698-651">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9a698-651">String &#124; Object</span></span>| | <span data-ttu-id="9a698-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9a698-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9a698-654">**OU**</span><span class="sxs-lookup"><span data-stu-id="9a698-654">**OR**</span></span><br/><span data-ttu-id="9a698-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="9a698-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9a698-657">String</span><span class="sxs-lookup"><span data-stu-id="9a698-657">String</span></span> | <span data-ttu-id="9a698-658">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-658">&lt;optional&gt;</span></span> | <span data-ttu-id="9a698-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="9a698-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="9a698-661">function</span><span class="sxs-lookup"><span data-stu-id="9a698-661">function</span></span> | <span data-ttu-id="9a698-662">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-662">&lt;optional&gt;</span></span> | <span data-ttu-id="9a698-663">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a698-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9a698-664">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-664">Requirements</span></span>

|<span data-ttu-id="9a698-665">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-665">Requirement</span></span>| <span data-ttu-id="9a698-666">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-667">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-668">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-668">1.0</span></span>|
|[<span data-ttu-id="9a698-669">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-670">ReadItem</span></span>|
|[<span data-ttu-id="9a698-671">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-672">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a698-673">Exemples</span><span class="sxs-lookup"><span data-stu-id="9a698-673">Examples</span></span>

<span data-ttu-id="9a698-674">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9a698-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9a698-675">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="9a698-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9a698-676">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="9a698-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9a698-677">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-677">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="9a698-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9a698-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="9a698-679">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9a698-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-680">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-681">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-681">Requirements</span></span>

|<span data-ttu-id="9a698-682">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-682">Requirement</span></span>| <span data-ttu-id="9a698-683">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-684">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-685">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-685">1.0</span></span>|
|[<span data-ttu-id="9a698-686">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-687">ReadItem</span></span>|
|[<span data-ttu-id="9a698-688">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-689">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a698-690">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9a698-690">Returns:</span></span>

<span data-ttu-id="9a698-691">Type : [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9a698-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9a698-692">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-692">Example</span></span>

<span data-ttu-id="9a698-693">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="9a698-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9a698-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9a698-695">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9a698-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-696">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-697">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-697">Parameters:</span></span>

|<span data-ttu-id="9a698-698">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-698">Name</span></span>| <span data-ttu-id="9a698-699">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-699">Type</span></span>| <span data-ttu-id="9a698-700">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9a698-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9a698-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="9a698-702">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="9a698-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a698-703">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-703">Requirements</span></span>

|<span data-ttu-id="9a698-704">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-704">Requirement</span></span>| <span data-ttu-id="9a698-705">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-706">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-707">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-707">1.0</span></span>|
|[<span data-ttu-id="9a698-708">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-709">Restreinte</span><span class="sxs-lookup"><span data-stu-id="9a698-709">Restricted</span></span>|
|[<span data-ttu-id="9a698-710">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-711">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a698-712">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9a698-712">Returns:</span></span>

<span data-ttu-id="9a698-713">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="9a698-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9a698-714">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="9a698-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9a698-715">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9a698-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9a698-716">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="9a698-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9a698-717">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="9a698-717">Value of `entityType`</span></span> | <span data-ttu-id="9a698-718">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="9a698-718">Type of objects in returned array</span></span> | <span data-ttu-id="9a698-719">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="9a698-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9a698-720">String</span><span class="sxs-lookup"><span data-stu-id="9a698-720">String</span></span> | <span data-ttu-id="9a698-721">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9a698-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9a698-722">Contact</span><span class="sxs-lookup"><span data-stu-id="9a698-722">Contact</span></span> | <span data-ttu-id="9a698-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a698-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9a698-724">String</span><span class="sxs-lookup"><span data-stu-id="9a698-724">String</span></span> | <span data-ttu-id="9a698-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a698-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9a698-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9a698-726">MeetingSuggestion</span></span> | <span data-ttu-id="9a698-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a698-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9a698-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9a698-728">PhoneNumber</span></span> | <span data-ttu-id="9a698-729">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9a698-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9a698-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9a698-730">TaskSuggestion</span></span> | <span data-ttu-id="9a698-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a698-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9a698-732">String</span><span class="sxs-lookup"><span data-stu-id="9a698-732">String</span></span> | <span data-ttu-id="9a698-733">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9a698-733">**Restricted**</span></span> |

<span data-ttu-id="9a698-734">Type :  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9a698-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="9a698-735">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-735">Example</span></span>

<span data-ttu-id="9a698-736">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="9a698-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="9a698-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9a698-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9a698-738">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9a698-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-739">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9a698-740">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="9a698-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-741">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-741">Parameters:</span></span>

|<span data-ttu-id="9a698-742">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-742">Name</span></span>| <span data-ttu-id="9a698-743">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-743">Type</span></span>| <span data-ttu-id="9a698-744">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9a698-745">String</span><span class="sxs-lookup"><span data-stu-id="9a698-745">String</span></span>|<span data-ttu-id="9a698-746">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="9a698-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a698-747">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-747">Requirements</span></span>

|<span data-ttu-id="9a698-748">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-748">Requirement</span></span>| <span data-ttu-id="9a698-749">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-750">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-751">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-751">1.0</span></span>|
|[<span data-ttu-id="9a698-752">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-753">ReadItem</span></span>|
|[<span data-ttu-id="9a698-754">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-755">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a698-756">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9a698-756">Returns:</span></span>

<span data-ttu-id="9a698-p146">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="9a698-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="9a698-759">Type : Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9a698-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="9a698-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9a698-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9a698-761">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9a698-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-762">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9a698-p147">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="9a698-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9a698-766">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="9a698-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9a698-767">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9a698-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="9a698-p148">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="9a698-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a698-770">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-770">Requirements</span></span>

|<span data-ttu-id="9a698-771">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-771">Requirement</span></span>| <span data-ttu-id="9a698-772">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-773">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-774">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-774">1.0</span></span>|
|[<span data-ttu-id="9a698-775">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-776">ReadItem</span></span>|
|[<span data-ttu-id="9a698-777">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-778">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a698-779">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9a698-779">Returns:</span></span>

<span data-ttu-id="9a698-p149">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="9a698-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9a698-782">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="9a698-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9a698-783">Object</span><span class="sxs-lookup"><span data-stu-id="9a698-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9a698-784">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-784">Example</span></span>

<span data-ttu-id="9a698-785">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="9a698-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9a698-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="9a698-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9a698-787">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9a698-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a698-788">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="9a698-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9a698-789">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="9a698-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9a698-p150">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="9a698-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-792">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-792">Parameters:</span></span>

|<span data-ttu-id="9a698-793">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-793">Name</span></span>| <span data-ttu-id="9a698-794">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-794">Type</span></span>| <span data-ttu-id="9a698-795">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9a698-796">String</span><span class="sxs-lookup"><span data-stu-id="9a698-796">String</span></span>|<span data-ttu-id="9a698-797">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="9a698-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a698-798">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-798">Requirements</span></span>

|<span data-ttu-id="9a698-799">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-799">Requirement</span></span>| <span data-ttu-id="9a698-800">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-801">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-802">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-802">1.0</span></span>|
|[<span data-ttu-id="9a698-803">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-804">ReadItem</span></span>|
|[<span data-ttu-id="9a698-805">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-806">Lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a698-807">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="9a698-807">Returns:</span></span>

<span data-ttu-id="9a698-808">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="9a698-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9a698-809">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="9a698-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9a698-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9a698-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9a698-811">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9a698-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9a698-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9a698-813">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="9a698-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9a698-p151">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="9a698-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-817">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-817">Parameters:</span></span>

|<span data-ttu-id="9a698-818">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-818">Name</span></span>| <span data-ttu-id="9a698-819">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-819">Type</span></span>| <span data-ttu-id="9a698-820">Attributs</span><span class="sxs-lookup"><span data-stu-id="9a698-820">Attributes</span></span>| <span data-ttu-id="9a698-821">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9a698-822">function</span><span class="sxs-lookup"><span data-stu-id="9a698-822">function</span></span>||<span data-ttu-id="9a698-823">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a698-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a698-824">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9a698-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9a698-825">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="9a698-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9a698-826">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-826">Object</span></span>| <span data-ttu-id="9a698-827">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-827">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-828">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9a698-829">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a698-830">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-830">Requirements</span></span>

|<span data-ttu-id="9a698-831">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-831">Requirement</span></span>| <span data-ttu-id="9a698-832">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-833">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-834">1.0</span><span class="sxs-lookup"><span data-stu-id="9a698-834">1.0</span></span>|
|[<span data-ttu-id="9a698-835">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a698-836">ReadItem</span></span>|
|[<span data-ttu-id="9a698-837">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-838">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a698-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-839">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-839">Example</span></span>

<span data-ttu-id="9a698-p154">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="9a698-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9a698-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a698-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9a698-844">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="9a698-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9a698-p155">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="9a698-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a698-849">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="9a698-849">Parameters:</span></span>

|<span data-ttu-id="9a698-850">Nom</span><span class="sxs-lookup"><span data-stu-id="9a698-850">Name</span></span>| <span data-ttu-id="9a698-851">Type</span><span class="sxs-lookup"><span data-stu-id="9a698-851">Type</span></span>| <span data-ttu-id="9a698-852">Attributs</span><span class="sxs-lookup"><span data-stu-id="9a698-852">Attributes</span></span>| <span data-ttu-id="9a698-853">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9a698-854">String</span><span class="sxs-lookup"><span data-stu-id="9a698-854">String</span></span>||<span data-ttu-id="9a698-855">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="9a698-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="9a698-856">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-856">Object</span></span>| <span data-ttu-id="9a698-857">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-857">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-858">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="9a698-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9a698-859">Objet</span><span class="sxs-lookup"><span data-stu-id="9a698-859">Object</span></span>| <span data-ttu-id="9a698-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-860">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-861">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9a698-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9a698-862">fonction</span><span class="sxs-lookup"><span data-stu-id="9a698-862">function</span></span>| <span data-ttu-id="9a698-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a698-863">&lt;optional&gt;</span></span>|<span data-ttu-id="9a698-864">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9a698-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a698-865">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="9a698-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a698-866">Erreurs</span><span class="sxs-lookup"><span data-stu-id="9a698-866">Errors</span></span>

| <span data-ttu-id="9a698-867">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="9a698-867">Error code</span></span> | <span data-ttu-id="9a698-868">Description</span><span class="sxs-lookup"><span data-stu-id="9a698-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9a698-869">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="9a698-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9a698-870">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a698-870">Requirements</span></span>

|<span data-ttu-id="9a698-871">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a698-871">Requirement</span></span>| <span data-ttu-id="9a698-872">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a698-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a698-873">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a698-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a698-874">1.1</span><span class="sxs-lookup"><span data-stu-id="9a698-874">1.1</span></span>|
|[<span data-ttu-id="9a698-875">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9a698-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a698-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a698-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a698-877">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a698-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9a698-878">Composition</span><span class="sxs-lookup"><span data-stu-id="9a698-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a698-879">Exemple</span><span class="sxs-lookup"><span data-stu-id="9a698-879">Example</span></span>

<span data-ttu-id="9a698-880">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="9a698-880">The following code removes an attachment with an identifier of '0'.</span></span>

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
