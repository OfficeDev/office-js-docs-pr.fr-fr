---
title: Office.Context.Mailbox.Item - exigence défini 1.1
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: ce8c10987c08609eba90a3a957b372114e62cd81
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701875"
---
# <a name="item"></a><span data-ttu-id="fe22a-102">élément</span><span class="sxs-lookup"><span data-stu-id="fe22a-102">item</span></span>

### <span data-ttu-id="fe22a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="fe22a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="fe22a-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="fe22a-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-107">Requirements</span></span>

|<span data-ttu-id="fe22a-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-108">Requirement</span></span>| <span data-ttu-id="fe22a-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-111">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-111">1.0</span></span>|
|[<span data-ttu-id="fe22a-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="fe22a-113">Restricted</span></span>|
|[<span data-ttu-id="fe22a-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-115">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="fe22a-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-116">Example</span></span>

<span data-ttu-id="fe22a-117">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="fe22a-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="fe22a-118">Membres</span><span class="sxs-lookup"><span data-stu-id="fe22a-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="fe22a-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fe22a-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="fe22a-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-122">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="fe22a-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fe22a-123">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="fe22a-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-124">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-124">Type:</span></span>

*   <span data-ttu-id="fe22a-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fe22a-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-126">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-126">Requirements</span></span>

|<span data-ttu-id="fe22a-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-127">Requirement</span></span>| <span data-ttu-id="fe22a-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-129">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-130">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-130">1.0</span></span>|
|[<span data-ttu-id="fe22a-131">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-132">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-134">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-135">Example</span></span>

<span data-ttu-id="fe22a-136">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="fe22a-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="fe22a-138">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="fe22a-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fe22a-139">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-140">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-140">Type:</span></span>

*   [<span data-ttu-id="fe22a-141">Destinataires</span><span class="sxs-lookup"><span data-stu-id="fe22a-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fe22a-142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-142">Requirements</span></span>

|<span data-ttu-id="fe22a-143">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-143">Requirement</span></span>| <span data-ttu-id="fe22a-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-145">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-146">1.1</span><span class="sxs-lookup"><span data-stu-id="fe22a-146">1.1</span></span>|
|[<span data-ttu-id="fe22a-147">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-148">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-149">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-150">Composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-151">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="fe22a-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="fe22a-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="fe22a-153">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="fe22a-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-154">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-154">Type:</span></span>

*   [<span data-ttu-id="fe22a-155">Corps</span><span class="sxs-lookup"><span data-stu-id="fe22a-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="fe22a-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-156">Requirements</span></span>

|<span data-ttu-id="fe22a-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-157">Requirement</span></span>| <span data-ttu-id="fe22a-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-160">1.1</span><span class="sxs-lookup"><span data-stu-id="fe22a-160">1.1</span></span>|
|[<span data-ttu-id="fe22a-161">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-162">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="fe22a-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="fe22a-166">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="fe22a-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fe22a-167">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-168">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-168">Read mode</span></span>

<span data-ttu-id="fe22a-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-171">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-171">Compose mode</span></span>

<span data-ttu-id="fe22a-172">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="fe22a-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-173">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-173">Type:</span></span>

*   <span data-ttu-id="fe22a-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-175">Requirements</span></span>

|<span data-ttu-id="fe22a-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-176">Requirement</span></span>| <span data-ttu-id="fe22a-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-179">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-179">1.0</span></span>|
|[<span data-ttu-id="fe22a-180">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-181">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-183">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-184">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fe22a-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="fe22a-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="fe22a-186">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="fe22a-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fe22a-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fe22a-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-191">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-191">Type:</span></span>

*   <span data-ttu-id="fe22a-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fe22a-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-193">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-193">Requirements</span></span>

|<span data-ttu-id="fe22a-194">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-194">Requirement</span></span>| <span data-ttu-id="fe22a-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-196">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-197">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-197">1.0</span></span>|
|[<span data-ttu-id="fe22a-198">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-199">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-201">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="fe22a-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="fe22a-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="fe22a-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-205">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-205">Type:</span></span>

*   <span data-ttu-id="fe22a-206">Date</span><span class="sxs-lookup"><span data-stu-id="fe22a-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-207">Requirements</span></span>

|<span data-ttu-id="fe22a-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-208">Requirement</span></span>| <span data-ttu-id="fe22a-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-211">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-211">1.0</span></span>|
|[<span data-ttu-id="fe22a-212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-213">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-215">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="fe22a-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="fe22a-217">dateTimeModified :Date</span></span>

<span data-ttu-id="fe22a-p111">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-220">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-221">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-221">Type:</span></span>

*   <span data-ttu-id="fe22a-222">Date</span><span class="sxs-lookup"><span data-stu-id="fe22a-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-223">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-223">Requirements</span></span>

|<span data-ttu-id="fe22a-224">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-224">Requirement</span></span>| <span data-ttu-id="fe22a-225">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-226">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-227">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-227">1.0</span></span>|
|[<span data-ttu-id="fe22a-228">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-229">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-230">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-231">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-232">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="fe22a-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe22a-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="fe22a-234">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fe22a-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-237">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-237">Read mode</span></span>

<span data-ttu-id="fe22a-238">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-239">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-239">Compose mode</span></span>

<span data-ttu-id="fe22a-240">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fe22a-241">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="fe22a-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-242">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-242">Type:</span></span>

*   <span data-ttu-id="fe22a-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe22a-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-244">Requirements</span></span>

|<span data-ttu-id="fe22a-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-245">Requirement</span></span>| <span data-ttu-id="fe22a-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-248">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-248">1.0</span></span>|
|[<span data-ttu-id="fe22a-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-250">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-252">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-253">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-253">Example</span></span>

<span data-ttu-id="fe22a-254">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="fe22a-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fe22a-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="fe22a-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="fe22a-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-260">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-261">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-261">Type:</span></span>

*   [<span data-ttu-id="fe22a-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fe22a-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fe22a-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-263">Requirements</span></span>

|<span data-ttu-id="fe22a-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-264">Requirement</span></span>| <span data-ttu-id="fe22a-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-267">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-267">1.0</span></span>|
|[<span data-ttu-id="fe22a-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-269">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-271">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="fe22a-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="fe22a-272">internetMessageId :String</span></span>

<span data-ttu-id="fe22a-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-275">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-275">Type:</span></span>

*   <span data-ttu-id="fe22a-276">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fe22a-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-277">Requirements</span></span>

|<span data-ttu-id="fe22a-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-278">Requirement</span></span>| <span data-ttu-id="fe22a-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-281">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-281">1.0</span></span>|
|[<span data-ttu-id="fe22a-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-283">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-285">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="fe22a-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="fe22a-287">itemClass :String</span></span>

<span data-ttu-id="fe22a-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fe22a-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="fe22a-292">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-292">Type</span></span> | <span data-ttu-id="fe22a-293">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-293">Description</span></span> | <span data-ttu-id="fe22a-294">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="fe22a-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="fe22a-295">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="fe22a-295">Appointment items</span></span> | <span data-ttu-id="fe22a-296">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="fe22a-297">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="fe22a-297">Message items</span></span> | <span data-ttu-id="fe22a-298">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="fe22a-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="fe22a-299">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-300">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-300">Type:</span></span>

*   <span data-ttu-id="fe22a-301">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fe22a-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-302">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-302">Requirements</span></span>

|<span data-ttu-id="fe22a-303">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-303">Requirement</span></span>| <span data-ttu-id="fe22a-304">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-305">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-306">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-306">1.0</span></span>|
|[<span data-ttu-id="fe22a-307">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-308">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-309">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-310">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-311">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fe22a-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="fe22a-312">(nullable) itemId :String</span></span>

<span data-ttu-id="fe22a-p118">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-315">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="fe22a-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fe22a-316">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="fe22a-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fe22a-317">Avant d’effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande `Office.context.mailbox.convertToRestId`, qui est disponible à partir de l’ensemble de conditions requises 1.3.</span><span class="sxs-lookup"><span data-stu-id="fe22a-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="fe22a-318">Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="fe22a-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-319">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-319">Type:</span></span>

*   <span data-ttu-id="fe22a-320">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fe22a-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-321">Requirements</span></span>

|<span data-ttu-id="fe22a-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-322">Requirement</span></span>| <span data-ttu-id="fe22a-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-325">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-325">1.0</span></span>|
|[<span data-ttu-id="fe22a-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-327">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-329">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-330">Example</span></span>

<span data-ttu-id="fe22a-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="fe22a-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fe22a-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fe22a-334">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="fe22a-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fe22a-335">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-336">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-336">Type:</span></span>

*   [<span data-ttu-id="fe22a-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fe22a-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fe22a-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-338">Requirements</span></span>

|<span data-ttu-id="fe22a-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-339">Requirement</span></span>| <span data-ttu-id="fe22a-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-342">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-342">1.0</span></span>|
|[<span data-ttu-id="fe22a-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-344">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-346">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="fe22a-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="fe22a-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="fe22a-349">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-350">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-350">Read mode</span></span>

<span data-ttu-id="fe22a-351">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-352">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-352">Compose mode</span></span>

<span data-ttu-id="fe22a-353">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-354">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-354">Type:</span></span>

*   <span data-ttu-id="fe22a-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="fe22a-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-356">Requirements</span></span>

|<span data-ttu-id="fe22a-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-357">Requirement</span></span>| <span data-ttu-id="fe22a-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-359">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-360">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-360">1.0</span></span>|
|[<span data-ttu-id="fe22a-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-362">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-364">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fe22a-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="fe22a-366">normalizedSubject :String</span></span>

<span data-ttu-id="fe22a-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fe22a-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="fe22a-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-371">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-371">Type:</span></span>

*   <span data-ttu-id="fe22a-372">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fe22a-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-373">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-373">Requirements</span></span>

|<span data-ttu-id="fe22a-374">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-374">Requirement</span></span>| <span data-ttu-id="fe22a-375">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-376">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-377">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-377">1.0</span></span>|
|[<span data-ttu-id="fe22a-378">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-379">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-380">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-381">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-382">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="fe22a-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="fe22a-384">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fe22a-385">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-386">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-386">Read mode</span></span>

<span data-ttu-id="fe22a-387">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="fe22a-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-388">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-388">Compose mode</span></span>

<span data-ttu-id="fe22a-389">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="fe22a-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-390">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-390">Type:</span></span>

*   <span data-ttu-id="fe22a-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-392">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-392">Requirements</span></span>

|<span data-ttu-id="fe22a-393">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-393">Requirement</span></span>| <span data-ttu-id="fe22a-394">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-395">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-396">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-396">1.0</span></span>|
|[<span data-ttu-id="fe22a-397">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-398">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-400">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-401">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="fe22a-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fe22a-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="fe22a-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-405">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-405">Type:</span></span>

*   [<span data-ttu-id="fe22a-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fe22a-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fe22a-407">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-407">Requirements</span></span>

|<span data-ttu-id="fe22a-408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-408">Requirement</span></span>| <span data-ttu-id="fe22a-409">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-410">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-411">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-411">1.0</span></span>|
|[<span data-ttu-id="fe22a-412">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-413">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-415">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-416">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="fe22a-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="fe22a-418">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fe22a-419">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-420">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-420">Read mode</span></span>

<span data-ttu-id="fe22a-421">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="fe22a-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-422">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-422">Compose mode</span></span>

<span data-ttu-id="fe22a-423">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="fe22a-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-424">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-424">Type:</span></span>

*   <span data-ttu-id="fe22a-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-426">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-426">Requirements</span></span>

|<span data-ttu-id="fe22a-427">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-427">Requirement</span></span>| <span data-ttu-id="fe22a-428">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-429">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-430">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-430">1.0</span></span>|
|[<span data-ttu-id="fe22a-431">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-432">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-433">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-434">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-435">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="fe22a-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fe22a-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="fe22a-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fe22a-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-441">La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-442">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-442">Type:</span></span>

*   [<span data-ttu-id="fe22a-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fe22a-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fe22a-444">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-444">Requirements</span></span>

|<span data-ttu-id="fe22a-445">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-445">Requirement</span></span>| <span data-ttu-id="fe22a-446">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-447">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-448">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-448">1.0</span></span>|
|[<span data-ttu-id="fe22a-449">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-450">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-451">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-452">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-453">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="fe22a-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe22a-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="fe22a-455">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fe22a-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-458">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-458">Read mode</span></span>

<span data-ttu-id="fe22a-459">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-460">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-460">Compose mode</span></span>

<span data-ttu-id="fe22a-461">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fe22a-462">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="fe22a-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-463">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-463">Type:</span></span>

*   <span data-ttu-id="fe22a-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe22a-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-465">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-465">Requirements</span></span>

|<span data-ttu-id="fe22a-466">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-466">Requirement</span></span>| <span data-ttu-id="fe22a-467">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-468">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-469">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-469">1.0</span></span>|
|[<span data-ttu-id="fe22a-470">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-471">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-472">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-473">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-474">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-474">Example</span></span>

<span data-ttu-id="fe22a-475">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="fe22a-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fe22a-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="fe22a-477">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="fe22a-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fe22a-478">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="fe22a-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-479">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-479">Read mode</span></span>

<span data-ttu-id="fe22a-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-482">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-482">Compose mode</span></span>

<span data-ttu-id="fe22a-483">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="fe22a-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fe22a-484">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-484">Type:</span></span>

*   <span data-ttu-id="fe22a-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fe22a-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-486">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-486">Requirements</span></span>

|<span data-ttu-id="fe22a-487">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-487">Requirement</span></span>| <span data-ttu-id="fe22a-488">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-489">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-490">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-490">1.0</span></span>|
|[<span data-ttu-id="fe22a-491">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-492">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-493">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-494">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="fe22a-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="fe22a-496">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="fe22a-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fe22a-497">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe22a-498">Mode Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-498">Read mode</span></span>

<span data-ttu-id="fe22a-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe22a-501">Mode composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-501">Compose mode</span></span>

<span data-ttu-id="fe22a-502">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="fe22a-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fe22a-503">Type :</span><span class="sxs-lookup"><span data-stu-id="fe22a-503">Type:</span></span>

*   <span data-ttu-id="fe22a-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe22a-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-505">Requirements</span></span>

|<span data-ttu-id="fe22a-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-506">Requirement</span></span>| <span data-ttu-id="fe22a-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-509">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-509">1.0</span></span>|
|[<span data-ttu-id="fe22a-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-511">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-513">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="fe22a-515">Méthodes</span><span class="sxs-lookup"><span data-stu-id="fe22a-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fe22a-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fe22a-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fe22a-517">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="fe22a-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fe22a-518">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="fe22a-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fe22a-519">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="fe22a-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-520">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-520">Parameters:</span></span>

|<span data-ttu-id="fe22a-521">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-521">Name</span></span>| <span data-ttu-id="fe22a-522">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-522">Type</span></span>| <span data-ttu-id="fe22a-523">Attributs</span><span class="sxs-lookup"><span data-stu-id="fe22a-523">Attributes</span></span>| <span data-ttu-id="fe22a-524">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="fe22a-525">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-525">String</span></span>||<span data-ttu-id="fe22a-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fe22a-528">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-528">String</span></span>||<span data-ttu-id="fe22a-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fe22a-531">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-531">Object</span></span>| <span data-ttu-id="fe22a-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-532">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-533">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe22a-534">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-534">Object</span></span>| <span data-ttu-id="fe22a-535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-535">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-536">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe22a-537">fonction</span><span class="sxs-lookup"><span data-stu-id="fe22a-537">function</span></span>| <span data-ttu-id="fe22a-538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-538">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-539">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe22a-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fe22a-540">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fe22a-541">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="fe22a-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fe22a-542">Erreurs</span><span class="sxs-lookup"><span data-stu-id="fe22a-542">Errors</span></span>

| <span data-ttu-id="fe22a-543">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="fe22a-543">Error code</span></span> | <span data-ttu-id="fe22a-544">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="fe22a-545">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="fe22a-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="fe22a-546">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="fe22a-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fe22a-547">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe22a-548">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-548">Requirements</span></span>

|<span data-ttu-id="fe22a-549">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-549">Requirement</span></span>| <span data-ttu-id="fe22a-550">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-551">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-552">1.1</span><span class="sxs-lookup"><span data-stu-id="fe22a-552">1.1</span></span>|
|[<span data-ttu-id="fe22a-553">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe22a-555">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-556">Composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-557">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fe22a-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fe22a-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fe22a-559">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fe22a-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fe22a-563">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="fe22a-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fe22a-564">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="fe22a-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-565">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-565">Parameters:</span></span>

|<span data-ttu-id="fe22a-566">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-566">Name</span></span>| <span data-ttu-id="fe22a-567">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-567">Type</span></span>| <span data-ttu-id="fe22a-568">Attributs</span><span class="sxs-lookup"><span data-stu-id="fe22a-568">Attributes</span></span>| <span data-ttu-id="fe22a-569">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="fe22a-570">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-570">String</span></span>||<span data-ttu-id="fe22a-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fe22a-573">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-573">String</span></span>||<span data-ttu-id="fe22a-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fe22a-576">Object</span><span class="sxs-lookup"><span data-stu-id="fe22a-576">Object</span></span>| <span data-ttu-id="fe22a-577">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-577">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-578">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe22a-579">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-579">Object</span></span>| <span data-ttu-id="fe22a-580">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-580">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-581">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe22a-582">fonction</span><span class="sxs-lookup"><span data-stu-id="fe22a-582">function</span></span>| <span data-ttu-id="fe22a-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-583">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-584">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe22a-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fe22a-585">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fe22a-586">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="fe22a-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fe22a-587">Erreurs</span><span class="sxs-lookup"><span data-stu-id="fe22a-587">Errors</span></span>

| <span data-ttu-id="fe22a-588">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="fe22a-588">Error code</span></span> | <span data-ttu-id="fe22a-589">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fe22a-590">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe22a-591">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-591">Requirements</span></span>

|<span data-ttu-id="fe22a-592">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-592">Requirement</span></span>| <span data-ttu-id="fe22a-593">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-594">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-595">1.1</span><span class="sxs-lookup"><span data-stu-id="fe22a-595">1.1</span></span>|
|[<span data-ttu-id="fe22a-596">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe22a-598">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-599">Composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-600">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-600">Example</span></span>

<span data-ttu-id="fe22a-601">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="fe22a-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fe22a-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="fe22a-603">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="fe22a-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-604">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe22a-605">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fe22a-606">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="fe22a-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-607">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyAllForm` n’est pas prise en charge dans l’ensemble des conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="fe22a-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="fe22a-608">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="fe22a-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-609">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-609">Parameters:</span></span>

|<span data-ttu-id="fe22a-610">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-610">Name</span></span>| <span data-ttu-id="fe22a-611">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-611">Type</span></span>| <span data-ttu-id="fe22a-612">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="fe22a-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fe22a-613">String &#124; Object</span></span>| |<span data-ttu-id="fe22a-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fe22a-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="fe22a-616">**OR**</span></span><br/><span data-ttu-id="fe22a-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="fe22a-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fe22a-619">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-619">String</span></span> | <span data-ttu-id="fe22a-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-620">&lt;optional&gt;</span></span> | <span data-ttu-id="fe22a-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="fe22a-623">function</span><span class="sxs-lookup"><span data-stu-id="fe22a-623">function</span></span> | <span data-ttu-id="fe22a-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-624">&lt;optional&gt;</span></span> | <span data-ttu-id="fe22a-625">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe22a-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe22a-626">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-626">Requirements</span></span>

|<span data-ttu-id="fe22a-627">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-627">Requirement</span></span>| <span data-ttu-id="fe22a-628">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-629">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-630">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-630">1.0</span></span>|
|[<span data-ttu-id="fe22a-631">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-632">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-633">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-634">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fe22a-635">Exemples</span><span class="sxs-lookup"><span data-stu-id="fe22a-635">Examples</span></span>

<span data-ttu-id="fe22a-636">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fe22a-637">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="fe22a-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fe22a-638">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="fe22a-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fe22a-639">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-639">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="fe22a-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fe22a-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="fe22a-641">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="fe22a-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-642">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe22a-643">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fe22a-644">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="fe22a-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-645">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyForm` n’est pas prise en charge dans l’ensemble des conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="fe22a-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="fe22a-646">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="fe22a-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-647">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-647">Parameters:</span></span>

|<span data-ttu-id="fe22a-648">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-648">Name</span></span>| <span data-ttu-id="fe22a-649">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-649">Type</span></span>| <span data-ttu-id="fe22a-650">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="fe22a-651">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fe22a-651">String &#124; Object</span></span>| | <span data-ttu-id="fe22a-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fe22a-654">**OU**</span><span class="sxs-lookup"><span data-stu-id="fe22a-654">**OR**</span></span><br/><span data-ttu-id="fe22a-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="fe22a-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fe22a-657">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-657">String</span></span> | <span data-ttu-id="fe22a-658">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-658">&lt;optional&gt;</span></span> | <span data-ttu-id="fe22a-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="fe22a-661">function</span><span class="sxs-lookup"><span data-stu-id="fe22a-661">function</span></span> | <span data-ttu-id="fe22a-662">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-662">&lt;optional&gt;</span></span> | <span data-ttu-id="fe22a-663">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe22a-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe22a-664">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-664">Requirements</span></span>

|<span data-ttu-id="fe22a-665">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-665">Requirement</span></span>| <span data-ttu-id="fe22a-666">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-667">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-668">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-668">1.0</span></span>|
|[<span data-ttu-id="fe22a-669">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-670">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-671">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-672">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fe22a-673">Exemples</span><span class="sxs-lookup"><span data-stu-id="fe22a-673">Examples</span></span>

<span data-ttu-id="fe22a-674">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fe22a-675">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="fe22a-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fe22a-676">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="fe22a-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fe22a-677">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-677">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="fe22a-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fe22a-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="fe22a-679">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="fe22a-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-680">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-681">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-681">Requirements</span></span>

|<span data-ttu-id="fe22a-682">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-682">Requirement</span></span>| <span data-ttu-id="fe22a-683">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-684">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-685">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-685">1.0</span></span>|
|[<span data-ttu-id="fe22a-686">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-687">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-688">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-689">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe22a-690">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="fe22a-690">Returns:</span></span>

<span data-ttu-id="fe22a-691">Type : [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fe22a-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fe22a-692">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-692">Example</span></span>

<span data-ttu-id="fe22a-693">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="fe22a-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fe22a-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fe22a-695">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="fe22a-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-696">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-697">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-697">Parameters:</span></span>

|<span data-ttu-id="fe22a-698">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-698">Name</span></span>| <span data-ttu-id="fe22a-699">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-699">Type</span></span>| <span data-ttu-id="fe22a-700">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="fe22a-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fe22a-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="fe22a-702">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="fe22a-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe22a-703">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-703">Requirements</span></span>

|<span data-ttu-id="fe22a-704">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-704">Requirement</span></span>| <span data-ttu-id="fe22a-705">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-706">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-707">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-707">1.0</span></span>|
|[<span data-ttu-id="fe22a-708">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-709">Restreinte</span><span class="sxs-lookup"><span data-stu-id="fe22a-709">Restricted</span></span>|
|[<span data-ttu-id="fe22a-710">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-711">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe22a-712">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="fe22a-712">Returns:</span></span>

<span data-ttu-id="fe22a-713">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="fe22a-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fe22a-714">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="fe22a-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="fe22a-715">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fe22a-716">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="fe22a-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="fe22a-717">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="fe22a-717">Value of `entityType`</span></span> | <span data-ttu-id="fe22a-718">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="fe22a-718">Type of objects in returned array</span></span> | <span data-ttu-id="fe22a-719">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="fe22a-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="fe22a-720">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-720">String</span></span> | <span data-ttu-id="fe22a-721">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fe22a-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="fe22a-722">Contact</span><span class="sxs-lookup"><span data-stu-id="fe22a-722">Contact</span></span> | <span data-ttu-id="fe22a-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe22a-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="fe22a-724">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-724">String</span></span> | <span data-ttu-id="fe22a-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe22a-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="fe22a-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fe22a-726">MeetingSuggestion</span></span> | <span data-ttu-id="fe22a-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe22a-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="fe22a-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fe22a-728">PhoneNumber</span></span> | <span data-ttu-id="fe22a-729">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fe22a-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="fe22a-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fe22a-730">TaskSuggestion</span></span> | <span data-ttu-id="fe22a-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe22a-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="fe22a-732">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-732">String</span></span> | <span data-ttu-id="fe22a-733">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fe22a-733">**Restricted**</span></span> |

<span data-ttu-id="fe22a-734">Type :  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fe22a-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="fe22a-735">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-735">Example</span></span>

<span data-ttu-id="fe22a-736">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="fe22a-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="fe22a-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fe22a-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fe22a-738">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fe22a-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-739">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe22a-740">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="fe22a-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-741">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-741">Parameters:</span></span>

|<span data-ttu-id="fe22a-742">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-742">Name</span></span>| <span data-ttu-id="fe22a-743">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-743">Type</span></span>| <span data-ttu-id="fe22a-744">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fe22a-745">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-745">String</span></span>|<span data-ttu-id="fe22a-746">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="fe22a-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe22a-747">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-747">Requirements</span></span>

|<span data-ttu-id="fe22a-748">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-748">Requirement</span></span>| <span data-ttu-id="fe22a-749">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-750">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-751">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-751">1.0</span></span>|
|[<span data-ttu-id="fe22a-752">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-753">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-754">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-755">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe22a-756">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="fe22a-756">Returns:</span></span>

<span data-ttu-id="fe22a-p146">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="fe22a-759">Type : Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fe22a-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="fe22a-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fe22a-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fe22a-761">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fe22a-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-762">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe22a-p147">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fe22a-766">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="fe22a-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fe22a-767">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="fe22a-p148">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe22a-770">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-770">Requirements</span></span>

|<span data-ttu-id="fe22a-771">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-771">Requirement</span></span>| <span data-ttu-id="fe22a-772">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-773">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-774">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-774">1.0</span></span>|
|[<span data-ttu-id="fe22a-775">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-776">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-777">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-778">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe22a-779">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="fe22a-779">Returns:</span></span>

<span data-ttu-id="fe22a-p149">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fe22a-782">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fe22a-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fe22a-783">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fe22a-784">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-784">Example</span></span>

<span data-ttu-id="fe22a-785">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="fe22a-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fe22a-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="fe22a-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fe22a-787">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fe22a-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fe22a-788">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="fe22a-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe22a-789">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="fe22a-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fe22a-p150">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-792">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-792">Parameters:</span></span>

|<span data-ttu-id="fe22a-793">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-793">Name</span></span>| <span data-ttu-id="fe22a-794">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-794">Type</span></span>| <span data-ttu-id="fe22a-795">object</span><span class="sxs-lookup"><span data-stu-id="fe22a-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fe22a-796">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-796">String</span></span>|<span data-ttu-id="fe22a-797">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="fe22a-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe22a-798">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-798">Requirements</span></span>

|<span data-ttu-id="fe22a-799">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-799">Requirement</span></span>| <span data-ttu-id="fe22a-800">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-801">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-802">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-802">1.0</span></span>|
|[<span data-ttu-id="fe22a-803">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-804">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-805">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-806">Lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe22a-807">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="fe22a-807">Returns:</span></span>

<span data-ttu-id="fe22a-808">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fe22a-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fe22a-809">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fe22a-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fe22a-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="fe22a-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fe22a-811">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fe22a-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fe22a-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fe22a-813">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="fe22a-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fe22a-p151">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-817">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-817">Parameters:</span></span>

|<span data-ttu-id="fe22a-818">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-818">Name</span></span>| <span data-ttu-id="fe22a-819">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-819">Type</span></span>| <span data-ttu-id="fe22a-820">Attributs</span><span class="sxs-lookup"><span data-stu-id="fe22a-820">Attributes</span></span>| <span data-ttu-id="fe22a-821">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fe22a-822">function</span><span class="sxs-lookup"><span data-stu-id="fe22a-822">function</span></span>||<span data-ttu-id="fe22a-823">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe22a-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fe22a-824">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe22a-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fe22a-825">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="fe22a-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="fe22a-826">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-826">Object</span></span>| <span data-ttu-id="fe22a-827">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-827">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-828">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="fe22a-829">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe22a-830">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-830">Requirements</span></span>

|<span data-ttu-id="fe22a-831">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-831">Requirement</span></span>| <span data-ttu-id="fe22a-832">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-833">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-834">1.0</span><span class="sxs-lookup"><span data-stu-id="fe22a-834">1.0</span></span>|
|[<span data-ttu-id="fe22a-835">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-836">ReadItem</span></span>|
|[<span data-ttu-id="fe22a-837">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-838">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fe22a-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-839">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-839">Example</span></span>

<span data-ttu-id="fe22a-p154">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fe22a-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fe22a-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fe22a-844">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="fe22a-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fe22a-p155">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="fe22a-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe22a-849">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="fe22a-849">Parameters:</span></span>

|<span data-ttu-id="fe22a-850">Nom</span><span class="sxs-lookup"><span data-stu-id="fe22a-850">Name</span></span>| <span data-ttu-id="fe22a-851">Type</span><span class="sxs-lookup"><span data-stu-id="fe22a-851">Type</span></span>| <span data-ttu-id="fe22a-852">Attributs</span><span class="sxs-lookup"><span data-stu-id="fe22a-852">Attributes</span></span>| <span data-ttu-id="fe22a-853">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="fe22a-854">String</span><span class="sxs-lookup"><span data-stu-id="fe22a-854">String</span></span>||<span data-ttu-id="fe22a-855">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="fe22a-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="fe22a-856">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-856">Object</span></span>| <span data-ttu-id="fe22a-857">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-857">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-858">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="fe22a-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe22a-859">Objet</span><span class="sxs-lookup"><span data-stu-id="fe22a-859">Object</span></span>| <span data-ttu-id="fe22a-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-860">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-861">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="fe22a-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe22a-862">fonction</span><span class="sxs-lookup"><span data-stu-id="fe22a-862">function</span></span>| <span data-ttu-id="fe22a-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe22a-863">&lt;optional&gt;</span></span>|<span data-ttu-id="fe22a-864">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe22a-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fe22a-865">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="fe22a-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fe22a-866">Erreurs</span><span class="sxs-lookup"><span data-stu-id="fe22a-866">Errors</span></span>

| <span data-ttu-id="fe22a-867">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="fe22a-867">Error code</span></span> | <span data-ttu-id="fe22a-868">Description</span><span class="sxs-lookup"><span data-stu-id="fe22a-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="fe22a-869">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="fe22a-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe22a-870">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fe22a-870">Requirements</span></span>

|<span data-ttu-id="fe22a-871">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fe22a-871">Requirement</span></span>| <span data-ttu-id="fe22a-872">Valeur</span><span class="sxs-lookup"><span data-stu-id="fe22a-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe22a-873">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fe22a-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe22a-874">1.1</span><span class="sxs-lookup"><span data-stu-id="fe22a-874">1.1</span></span>|
|[<span data-ttu-id="fe22a-875">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fe22a-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe22a-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe22a-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe22a-877">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fe22a-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe22a-878">Composition</span><span class="sxs-lookup"><span data-stu-id="fe22a-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe22a-879">Exemple</span><span class="sxs-lookup"><span data-stu-id="fe22a-879">Example</span></span>

<span data-ttu-id="fe22a-880">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="fe22a-880">The following code removes an attachment with an identifier of '0'.</span></span>

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
