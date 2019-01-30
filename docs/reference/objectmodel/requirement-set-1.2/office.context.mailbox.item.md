---
title: Office.Context.Mailbox.Item - exigence défini 1.2
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: d58a38ce045a179a7e5cdd2e15b4e16c2ac03c91
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388597"
---
# <a name="item"></a><span data-ttu-id="046f8-102">élément</span><span class="sxs-lookup"><span data-stu-id="046f8-102">item</span></span>

### <span data-ttu-id="046f8-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="046f8-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="046f8-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="046f8-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-107">Requirements</span></span>

|<span data-ttu-id="046f8-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-108">Requirement</span></span>| <span data-ttu-id="046f8-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-111">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-111">1.0</span></span>|
|[<span data-ttu-id="046f8-112">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-113">Restreinte</span><span class="sxs-lookup"><span data-stu-id="046f8-113">Restricted</span></span>|
|[<span data-ttu-id="046f8-114">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-115">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="046f8-116">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-116">Example</span></span>

<span data-ttu-id="046f8-117">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="046f8-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="046f8-118">Membres</span><span class="sxs-lookup"><span data-stu-id="046f8-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="046f8-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="046f8-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="046f8-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-122">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="046f8-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="046f8-123">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="046f8-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-124">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-124">Type:</span></span>

*   <span data-ttu-id="046f8-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="046f8-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-126">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-126">Requirements</span></span>

|<span data-ttu-id="046f8-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-127">Requirement</span></span>| <span data-ttu-id="046f8-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-129">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-130">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-130">1.0</span></span>|
|[<span data-ttu-id="046f8-131">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-132">ReadItem</span></span>|
|[<span data-ttu-id="046f8-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-134">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-135">Example</span></span>

<span data-ttu-id="046f8-136">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="046f8-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="046f8-138">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="046f8-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="046f8-139">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-140">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-140">Type:</span></span>

*   [<span data-ttu-id="046f8-141">Destinataires</span><span class="sxs-lookup"><span data-stu-id="046f8-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="046f8-142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-142">Requirements</span></span>

|<span data-ttu-id="046f8-143">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-143">Requirement</span></span>| <span data-ttu-id="046f8-144">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-145">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-146">1.1</span><span class="sxs-lookup"><span data-stu-id="046f8-146">1.1</span></span>|
|[<span data-ttu-id="046f8-147">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-148">ReadItem</span></span>|
|[<span data-ttu-id="046f8-149">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-150">Composition</span><span class="sxs-lookup"><span data-stu-id="046f8-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-151">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="046f8-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="046f8-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="046f8-153">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-154">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-154">Type:</span></span>

*   [<span data-ttu-id="046f8-155">Corps</span><span class="sxs-lookup"><span data-stu-id="046f8-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="046f8-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-156">Requirements</span></span>

|<span data-ttu-id="046f8-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-157">Requirement</span></span>| <span data-ttu-id="046f8-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-160">1.1</span><span class="sxs-lookup"><span data-stu-id="046f8-160">1.1</span></span>|
|[<span data-ttu-id="046f8-161">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-162">ReadItem</span></span>|
|[<span data-ttu-id="046f8-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="046f8-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="046f8-166">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="046f8-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="046f8-167">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-168">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-168">Read mode</span></span>

<span data-ttu-id="046f8-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="046f8-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-171">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-171">Compose mode</span></span>

<span data-ttu-id="046f8-172">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="046f8-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-173">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-173">Type:</span></span>

*   <span data-ttu-id="046f8-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-175">Requirements</span></span>

|<span data-ttu-id="046f8-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-176">Requirement</span></span>| <span data-ttu-id="046f8-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-179">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-179">1.0</span></span>|
|[<span data-ttu-id="046f8-180">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-181">ReadItem</span></span>|
|[<span data-ttu-id="046f8-182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-183">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-184">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="046f8-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="046f8-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="046f8-186">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="046f8-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="046f8-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="046f8-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="046f8-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="046f8-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-191">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-191">Type:</span></span>

*   <span data-ttu-id="046f8-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-193">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-193">Requirements</span></span>

|<span data-ttu-id="046f8-194">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-194">Requirement</span></span>| <span data-ttu-id="046f8-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-196">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-197">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-197">1.0</span></span>|
|[<span data-ttu-id="046f8-198">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-199">ReadItem</span></span>|
|[<span data-ttu-id="046f8-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-201">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="046f8-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="046f8-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="046f8-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-205">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-205">Type:</span></span>

*   <span data-ttu-id="046f8-206">Date</span><span class="sxs-lookup"><span data-stu-id="046f8-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-207">Requirements</span></span>

|<span data-ttu-id="046f8-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-208">Requirement</span></span>| <span data-ttu-id="046f8-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-211">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-211">1.0</span></span>|
|[<span data-ttu-id="046f8-212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-213">ReadItem</span></span>|
|[<span data-ttu-id="046f8-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-215">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="046f8-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="046f8-217">dateTimeModified :Date</span></span>

<span data-ttu-id="046f8-p111">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-220">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-221">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-221">Type:</span></span>

*   <span data-ttu-id="046f8-222">Date</span><span class="sxs-lookup"><span data-stu-id="046f8-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-223">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-223">Requirements</span></span>

|<span data-ttu-id="046f8-224">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-224">Requirement</span></span>| <span data-ttu-id="046f8-225">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-226">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-227">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-227">1.0</span></span>|
|[<span data-ttu-id="046f8-228">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-229">ReadItem</span></span>|
|[<span data-ttu-id="046f8-230">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-231">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-232">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="046f8-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="046f8-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="046f8-234">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="046f8-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="046f8-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-237">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-237">Read mode</span></span>

<span data-ttu-id="046f8-238">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="046f8-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-239">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-239">Compose mode</span></span>

<span data-ttu-id="046f8-240">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="046f8-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="046f8-241">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="046f8-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-242">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-242">Type:</span></span>

*   <span data-ttu-id="046f8-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="046f8-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-244">Requirements</span></span>

|<span data-ttu-id="046f8-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-245">Requirement</span></span>| <span data-ttu-id="046f8-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-248">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-248">1.0</span></span>|
|[<span data-ttu-id="046f8-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-250">ReadItem</span></span>|
|[<span data-ttu-id="046f8-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-252">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-253">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-253">Example</span></span>

<span data-ttu-id="046f8-254">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="046f8-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="046f8-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="046f8-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="046f8-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="046f8-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="046f8-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-260">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="046f8-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-261">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-261">Type:</span></span>

*   [<span data-ttu-id="046f8-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="046f8-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="046f8-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-263">Requirements</span></span>

|<span data-ttu-id="046f8-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-264">Requirement</span></span>| <span data-ttu-id="046f8-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-267">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-267">1.0</span></span>|
|[<span data-ttu-id="046f8-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-269">ReadItem</span></span>|
|[<span data-ttu-id="046f8-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-271">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="046f8-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="046f8-272">internetMessageId :String</span></span>

<span data-ttu-id="046f8-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-275">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-275">Type:</span></span>

*   <span data-ttu-id="046f8-276">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-277">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-277">Requirements</span></span>

|<span data-ttu-id="046f8-278">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-278">Requirement</span></span>| <span data-ttu-id="046f8-279">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-280">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-281">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-281">1.0</span></span>|
|[<span data-ttu-id="046f8-282">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-283">ReadItem</span></span>|
|[<span data-ttu-id="046f8-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-285">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-286">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="046f8-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="046f8-287">itemClass :String</span></span>

<span data-ttu-id="046f8-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="046f8-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="046f8-292">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-292">Type</span></span> | <span data-ttu-id="046f8-293">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-293">Description</span></span> | <span data-ttu-id="046f8-294">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="046f8-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="046f8-295">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="046f8-295">Appointment items</span></span> | <span data-ttu-id="046f8-296">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="046f8-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="046f8-297">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="046f8-297">Message items</span></span> | <span data-ttu-id="046f8-298">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="046f8-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="046f8-299">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="046f8-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-300">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-300">Type:</span></span>

*   <span data-ttu-id="046f8-301">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-302">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-302">Requirements</span></span>

|<span data-ttu-id="046f8-303">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-303">Requirement</span></span>| <span data-ttu-id="046f8-304">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-305">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-306">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-306">1.0</span></span>|
|[<span data-ttu-id="046f8-307">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-308">ReadItem</span></span>|
|[<span data-ttu-id="046f8-309">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-310">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-311">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="046f8-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="046f8-312">(nullable) itemId :String</span></span>

<span data-ttu-id="046f8-p118">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-315">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="046f8-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="046f8-316">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="046f8-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="046f8-317">Avant d’effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande `Office.context.mailbox.convertToRestId`, qui est disponible à partir de l’ensemble de conditions requises 1.3.</span><span class="sxs-lookup"><span data-stu-id="046f8-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="046f8-318">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="046f8-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-319">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-319">Type:</span></span>

*   <span data-ttu-id="046f8-320">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-321">Requirements</span></span>

|<span data-ttu-id="046f8-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-322">Requirement</span></span>| <span data-ttu-id="046f8-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-325">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-325">1.0</span></span>|
|[<span data-ttu-id="046f8-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-327">ReadItem</span></span>|
|[<span data-ttu-id="046f8-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-329">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-330">Example</span></span>

<span data-ttu-id="046f8-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="046f8-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="046f8-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="046f8-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="046f8-334">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="046f8-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="046f8-335">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-336">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-336">Type:</span></span>

*   [<span data-ttu-id="046f8-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="046f8-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="046f8-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-338">Requirements</span></span>

|<span data-ttu-id="046f8-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-339">Requirement</span></span>| <span data-ttu-id="046f8-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-342">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-342">1.0</span></span>|
|[<span data-ttu-id="046f8-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-344">ReadItem</span></span>|
|[<span data-ttu-id="046f8-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-346">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="046f8-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="046f8-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="046f8-349">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-350">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-350">Read mode</span></span>

<span data-ttu-id="046f8-351">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-352">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-352">Compose mode</span></span>

<span data-ttu-id="046f8-353">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-354">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-354">Type:</span></span>

*   <span data-ttu-id="046f8-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="046f8-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-356">Requirements</span></span>

|<span data-ttu-id="046f8-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-357">Requirement</span></span>| <span data-ttu-id="046f8-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-359">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-360">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-360">1.0</span></span>|
|[<span data-ttu-id="046f8-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-362">ReadItem</span></span>|
|[<span data-ttu-id="046f8-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-364">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="046f8-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="046f8-366">normalizedSubject :String</span></span>

<span data-ttu-id="046f8-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="046f8-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="046f8-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-371">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-371">Type:</span></span>

*   <span data-ttu-id="046f8-372">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-373">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-373">Requirements</span></span>

|<span data-ttu-id="046f8-374">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-374">Requirement</span></span>| <span data-ttu-id="046f8-375">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-376">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-377">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-377">1.0</span></span>|
|[<span data-ttu-id="046f8-378">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-379">ReadItem</span></span>|
|[<span data-ttu-id="046f8-380">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-381">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-382">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="046f8-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="046f8-384">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="046f8-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="046f8-385">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-386">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-386">Read mode</span></span>

<span data-ttu-id="046f8-387">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="046f8-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-388">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-388">Compose mode</span></span>

<span data-ttu-id="046f8-389">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="046f8-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-390">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-390">Type:</span></span>

*   <span data-ttu-id="046f8-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-392">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-392">Requirements</span></span>

|<span data-ttu-id="046f8-393">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-393">Requirement</span></span>| <span data-ttu-id="046f8-394">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-395">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-396">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-396">1.0</span></span>|
|[<span data-ttu-id="046f8-397">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-398">ReadItem</span></span>|
|[<span data-ttu-id="046f8-399">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-400">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-401">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="046f8-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="046f8-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="046f8-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-405">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-405">Type:</span></span>

*   [<span data-ttu-id="046f8-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="046f8-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="046f8-407">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-407">Requirements</span></span>

|<span data-ttu-id="046f8-408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-408">Requirement</span></span>| <span data-ttu-id="046f8-409">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-410">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-411">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-411">1.0</span></span>|
|[<span data-ttu-id="046f8-412">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-413">ReadItem</span></span>|
|[<span data-ttu-id="046f8-414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-415">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-416">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="046f8-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="046f8-418">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="046f8-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="046f8-419">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-420">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-420">Read mode</span></span>

<span data-ttu-id="046f8-421">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="046f8-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-422">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-422">Compose mode</span></span>

<span data-ttu-id="046f8-423">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="046f8-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-424">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-424">Type:</span></span>

*   <span data-ttu-id="046f8-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-426">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-426">Requirements</span></span>

|<span data-ttu-id="046f8-427">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-427">Requirement</span></span>| <span data-ttu-id="046f8-428">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-429">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-430">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-430">1.0</span></span>|
|[<span data-ttu-id="046f8-431">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-432">ReadItem</span></span>|
|[<span data-ttu-id="046f8-433">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-434">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-435">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="046f8-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="046f8-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="046f8-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="046f8-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="046f8-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="046f8-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-441">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="046f8-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-442">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-442">Type:</span></span>

*   [<span data-ttu-id="046f8-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="046f8-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="046f8-444">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-444">Requirements</span></span>

|<span data-ttu-id="046f8-445">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-445">Requirement</span></span>| <span data-ttu-id="046f8-446">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-447">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-448">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-448">1.0</span></span>|
|[<span data-ttu-id="046f8-449">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-450">ReadItem</span></span>|
|[<span data-ttu-id="046f8-451">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-452">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-453">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="046f8-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="046f8-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="046f8-455">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="046f8-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="046f8-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-458">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-458">Read mode</span></span>

<span data-ttu-id="046f8-459">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="046f8-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-460">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-460">Compose mode</span></span>

<span data-ttu-id="046f8-461">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="046f8-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="046f8-462">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="046f8-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-463">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-463">Type:</span></span>

*   <span data-ttu-id="046f8-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="046f8-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-465">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-465">Requirements</span></span>

|<span data-ttu-id="046f8-466">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-466">Requirement</span></span>| <span data-ttu-id="046f8-467">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-468">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-469">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-469">1.0</span></span>|
|[<span data-ttu-id="046f8-470">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-471">ReadItem</span></span>|
|[<span data-ttu-id="046f8-472">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-473">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-474">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-474">Example</span></span>

<span data-ttu-id="046f8-475">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="046f8-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="046f8-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="046f8-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="046f8-477">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="046f8-478">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="046f8-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-479">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-479">Read mode</span></span>

<span data-ttu-id="046f8-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="046f8-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="046f8-482">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-482">Compose mode</span></span>

<span data-ttu-id="046f8-483">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="046f8-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="046f8-484">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-484">Type:</span></span>

*   <span data-ttu-id="046f8-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="046f8-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-486">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-486">Requirements</span></span>

|<span data-ttu-id="046f8-487">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-487">Requirement</span></span>| <span data-ttu-id="046f8-488">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-489">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-490">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-490">1.0</span></span>|
|[<span data-ttu-id="046f8-491">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-492">ReadItem</span></span>|
|[<span data-ttu-id="046f8-493">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-494">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="046f8-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="046f8-496">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="046f8-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="046f8-497">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="046f8-498">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-498">Read mode</span></span>

<span data-ttu-id="046f8-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="046f8-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="046f8-501">Mode composition</span><span class="sxs-lookup"><span data-stu-id="046f8-501">Compose mode</span></span>

<span data-ttu-id="046f8-502">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="046f8-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="046f8-503">Type :</span><span class="sxs-lookup"><span data-stu-id="046f8-503">Type:</span></span>

*   <span data-ttu-id="046f8-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="046f8-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-505">Requirements</span></span>

|<span data-ttu-id="046f8-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-506">Requirement</span></span>| <span data-ttu-id="046f8-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-509">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-509">1.0</span></span>|
|[<span data-ttu-id="046f8-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-511">ReadItem</span></span>|
|[<span data-ttu-id="046f8-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-513">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="046f8-515">Méthodes</span><span class="sxs-lookup"><span data-stu-id="046f8-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="046f8-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="046f8-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="046f8-517">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="046f8-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="046f8-518">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="046f8-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="046f8-519">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="046f8-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-520">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-520">Parameters:</span></span>

|<span data-ttu-id="046f8-521">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-521">Name</span></span>| <span data-ttu-id="046f8-522">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-522">Type</span></span>| <span data-ttu-id="046f8-523">Attributs</span><span class="sxs-lookup"><span data-stu-id="046f8-523">Attributes</span></span>| <span data-ttu-id="046f8-524">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="046f8-525">String</span><span class="sxs-lookup"><span data-stu-id="046f8-525">String</span></span>||<span data-ttu-id="046f8-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="046f8-528">String</span><span class="sxs-lookup"><span data-stu-id="046f8-528">String</span></span>||<span data-ttu-id="046f8-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="046f8-531">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-531">Object</span></span>| <span data-ttu-id="046f8-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-532">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-533">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="046f8-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="046f8-534">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-534">Object</span></span>| <span data-ttu-id="046f8-535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-535">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-536">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="046f8-537">fonction</span><span class="sxs-lookup"><span data-stu-id="046f8-537">function</span></span>| <span data-ttu-id="046f8-538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-538">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-539">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="046f8-540">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="046f8-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="046f8-541">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="046f8-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="046f8-542">Erreurs</span><span class="sxs-lookup"><span data-stu-id="046f8-542">Errors</span></span>

| <span data-ttu-id="046f8-543">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="046f8-543">Error code</span></span> | <span data-ttu-id="046f8-544">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="046f8-545">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="046f8-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="046f8-546">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="046f8-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="046f8-547">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="046f8-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="046f8-548">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-548">Requirements</span></span>

|<span data-ttu-id="046f8-549">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-549">Requirement</span></span>| <span data-ttu-id="046f8-550">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-551">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-552">1.1</span><span class="sxs-lookup"><span data-stu-id="046f8-552">1.1</span></span>|
|[<span data-ttu-id="046f8-553">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="046f8-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="046f8-555">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-556">Composition</span><span class="sxs-lookup"><span data-stu-id="046f8-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-557">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="046f8-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="046f8-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="046f8-559">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="046f8-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="046f8-563">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="046f8-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="046f8-564">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="046f8-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-565">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-565">Parameters:</span></span>

|<span data-ttu-id="046f8-566">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-566">Name</span></span>| <span data-ttu-id="046f8-567">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-567">Type</span></span>| <span data-ttu-id="046f8-568">Attributs</span><span class="sxs-lookup"><span data-stu-id="046f8-568">Attributes</span></span>| <span data-ttu-id="046f8-569">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="046f8-570">String</span><span class="sxs-lookup"><span data-stu-id="046f8-570">String</span></span>||<span data-ttu-id="046f8-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="046f8-573">String</span><span class="sxs-lookup"><span data-stu-id="046f8-573">String</span></span>||<span data-ttu-id="046f8-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="046f8-576">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-576">Object</span></span>| <span data-ttu-id="046f8-577">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-577">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-578">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="046f8-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="046f8-579">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-579">Object</span></span>| <span data-ttu-id="046f8-580">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-580">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-581">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="046f8-582">fonction</span><span class="sxs-lookup"><span data-stu-id="046f8-582">function</span></span>| <span data-ttu-id="046f8-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-583">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-584">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="046f8-585">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="046f8-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="046f8-586">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="046f8-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="046f8-587">Erreurs</span><span class="sxs-lookup"><span data-stu-id="046f8-587">Errors</span></span>

| <span data-ttu-id="046f8-588">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="046f8-588">Error code</span></span> | <span data-ttu-id="046f8-589">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="046f8-590">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="046f8-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="046f8-591">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-591">Requirements</span></span>

|<span data-ttu-id="046f8-592">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-592">Requirement</span></span>| <span data-ttu-id="046f8-593">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-594">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-595">1.1</span><span class="sxs-lookup"><span data-stu-id="046f8-595">1.1</span></span>|
|[<span data-ttu-id="046f8-596">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="046f8-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="046f8-598">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-599">Composition</span><span class="sxs-lookup"><span data-stu-id="046f8-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-600">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-600">Example</span></span>

<span data-ttu-id="046f8-601">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="046f8-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="046f8-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="046f8-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="046f8-603">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="046f8-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-604">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="046f8-605">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="046f8-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="046f8-606">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="046f8-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="046f8-p137">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="046f8-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-610">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-610">Parameters:</span></span>

|<span data-ttu-id="046f8-611">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-611">Name</span></span>| <span data-ttu-id="046f8-612">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-612">Type</span></span>| <span data-ttu-id="046f8-613">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-613">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="046f8-614">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="046f8-614">String &#124; Object</span></span>| |<span data-ttu-id="046f8-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="046f8-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="046f8-617">**OU**</span><span class="sxs-lookup"><span data-stu-id="046f8-617">**OR**</span></span><br/><span data-ttu-id="046f8-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="046f8-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="046f8-620">String</span><span class="sxs-lookup"><span data-stu-id="046f8-620">String</span></span> | <span data-ttu-id="046f8-621">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-621">&lt;optional&gt;</span></span> | <span data-ttu-id="046f8-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="046f8-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="046f8-624">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-624">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="046f8-625">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-625">&lt;optional&gt;</span></span> | <span data-ttu-id="046f8-626">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-626">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="046f8-627">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-627">String</span></span> | | <span data-ttu-id="046f8-p141">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="046f8-630">String</span><span class="sxs-lookup"><span data-stu-id="046f8-630">String</span></span> | | <span data-ttu-id="046f8-631">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-631">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="046f8-632">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-632">String</span></span> | | <span data-ttu-id="046f8-p142">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="046f8-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="046f8-635">String</span><span class="sxs-lookup"><span data-stu-id="046f8-635">String</span></span> | | <span data-ttu-id="046f8-p143">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="046f8-639">function</span><span class="sxs-lookup"><span data-stu-id="046f8-639">function</span></span> | <span data-ttu-id="046f8-640">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-640">&lt;optional&gt;</span></span> | <span data-ttu-id="046f8-641">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-641">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="046f8-642">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-642">Requirements</span></span>

|<span data-ttu-id="046f8-643">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-643">Requirement</span></span>| <span data-ttu-id="046f8-644">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-645">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-646">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-646">1.0</span></span>|
|[<span data-ttu-id="046f8-647">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-648">ReadItem</span></span>|
|[<span data-ttu-id="046f8-649">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-650">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-650">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="046f8-651">Exemples</span><span class="sxs-lookup"><span data-stu-id="046f8-651">Examples</span></span>

<span data-ttu-id="046f8-652">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="046f8-652">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="046f8-653">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="046f8-653">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="046f8-654">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="046f8-654">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="046f8-655">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="046f8-655">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="046f8-656">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-656">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="046f8-657">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-657">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="046f8-658">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="046f8-658">displayReplyForm(formData)</span></span>

<span data-ttu-id="046f8-659">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="046f8-659">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-660">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-660">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="046f8-661">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="046f8-661">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="046f8-662">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="046f8-662">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="046f8-p144">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="046f8-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-666">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-666">Parameters:</span></span>

|<span data-ttu-id="046f8-667">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-667">Name</span></span>| <span data-ttu-id="046f8-668">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-668">Type</span></span>| <span data-ttu-id="046f8-669">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-669">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="046f8-670">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="046f8-670">String &#124; Object</span></span>| | <span data-ttu-id="046f8-p145">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="046f8-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="046f8-673">**OU**</span><span class="sxs-lookup"><span data-stu-id="046f8-673">**OR**</span></span><br/><span data-ttu-id="046f8-p146">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="046f8-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="046f8-676">String</span><span class="sxs-lookup"><span data-stu-id="046f8-676">String</span></span> | <span data-ttu-id="046f8-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-677">&lt;optional&gt;</span></span> | <span data-ttu-id="046f8-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="046f8-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="046f8-680">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-680">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="046f8-681">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-681">&lt;optional&gt;</span></span> | <span data-ttu-id="046f8-682">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-682">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="046f8-683">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-683">String</span></span> | | <span data-ttu-id="046f8-p148">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="046f8-686">String</span><span class="sxs-lookup"><span data-stu-id="046f8-686">String</span></span> | | <span data-ttu-id="046f8-687">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-687">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="046f8-688">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-688">String</span></span> | | <span data-ttu-id="046f8-p149">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="046f8-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="046f8-691">String</span><span class="sxs-lookup"><span data-stu-id="046f8-691">String</span></span> | | <span data-ttu-id="046f8-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="046f8-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="046f8-695">function</span><span class="sxs-lookup"><span data-stu-id="046f8-695">function</span></span> | <span data-ttu-id="046f8-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-696">&lt;optional&gt;</span></span> | <span data-ttu-id="046f8-697">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="046f8-698">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-698">Requirements</span></span>

|<span data-ttu-id="046f8-699">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-699">Requirement</span></span>| <span data-ttu-id="046f8-700">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-700">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-701">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-701">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-702">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-702">1.0</span></span>|
|[<span data-ttu-id="046f8-703">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-703">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-704">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-704">ReadItem</span></span>|
|[<span data-ttu-id="046f8-705">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-705">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-706">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-706">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="046f8-707">Exemples</span><span class="sxs-lookup"><span data-stu-id="046f8-707">Examples</span></span>

<span data-ttu-id="046f8-708">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="046f8-708">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="046f8-709">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="046f8-709">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="046f8-710">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="046f8-710">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="046f8-711">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="046f8-711">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="046f8-712">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="046f8-712">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="046f8-713">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-713">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="046f8-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="046f8-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="046f8-715">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="046f8-715">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-716">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-717">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-717">Requirements</span></span>

|<span data-ttu-id="046f8-718">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-718">Requirement</span></span>| <span data-ttu-id="046f8-719">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-720">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-721">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-721">1.0</span></span>|
|[<span data-ttu-id="046f8-722">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-722">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-723">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-723">ReadItem</span></span>|
|[<span data-ttu-id="046f8-724">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-724">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-725">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-725">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="046f8-726">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="046f8-726">Returns:</span></span>

<span data-ttu-id="046f8-727">Type : [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="046f8-727">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="046f8-728">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-728">Example</span></span>

<span data-ttu-id="046f8-729">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-729">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="046f8-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="046f8-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="046f8-731">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="046f8-731">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-732">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-732">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-733">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-733">Parameters:</span></span>

|<span data-ttu-id="046f8-734">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-734">Name</span></span>| <span data-ttu-id="046f8-735">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-735">Type</span></span>| <span data-ttu-id="046f8-736">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-736">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="046f8-737">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="046f8-737">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="046f8-738">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="046f8-738">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="046f8-739">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-739">Requirements</span></span>

|<span data-ttu-id="046f8-740">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-740">Requirement</span></span>| <span data-ttu-id="046f8-741">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-742">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-743">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-743">1.0</span></span>|
|[<span data-ttu-id="046f8-744">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-744">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-745">Restreinte</span><span class="sxs-lookup"><span data-stu-id="046f8-745">Restricted</span></span>|
|[<span data-ttu-id="046f8-746">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-746">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-747">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-747">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="046f8-748">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="046f8-748">Returns:</span></span>

<span data-ttu-id="046f8-749">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="046f8-749">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="046f8-750">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="046f8-750">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="046f8-751">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="046f8-751">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="046f8-752">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="046f8-752">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="046f8-753">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="046f8-753">Value of `entityType`</span></span> | <span data-ttu-id="046f8-754">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="046f8-754">Type of objects in returned array</span></span> | <span data-ttu-id="046f8-755">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="046f8-755">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="046f8-756">String</span><span class="sxs-lookup"><span data-stu-id="046f8-756">String</span></span> | <span data-ttu-id="046f8-757">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="046f8-757">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="046f8-758">Contact</span><span class="sxs-lookup"><span data-stu-id="046f8-758">Contact</span></span> | <span data-ttu-id="046f8-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="046f8-759">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="046f8-760">String</span><span class="sxs-lookup"><span data-stu-id="046f8-760">String</span></span> | <span data-ttu-id="046f8-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="046f8-761">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="046f8-762">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="046f8-762">MeetingSuggestion</span></span> | <span data-ttu-id="046f8-763">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="046f8-763">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="046f8-764">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="046f8-764">PhoneNumber</span></span> | <span data-ttu-id="046f8-765">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="046f8-765">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="046f8-766">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="046f8-766">TaskSuggestion</span></span> | <span data-ttu-id="046f8-767">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="046f8-767">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="046f8-768">String</span><span class="sxs-lookup"><span data-stu-id="046f8-768">String</span></span> | <span data-ttu-id="046f8-769">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="046f8-769">**Restricted**</span></span> |

<span data-ttu-id="046f8-770">Type : Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="046f8-770">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="046f8-771">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-771">Example</span></span>

<span data-ttu-id="046f8-772">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="046f8-772">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="046f8-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="046f8-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="046f8-774">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="046f8-774">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-775">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="046f8-776">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="046f8-776">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-777">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-777">Parameters:</span></span>

|<span data-ttu-id="046f8-778">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-778">Name</span></span>| <span data-ttu-id="046f8-779">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-779">Type</span></span>| <span data-ttu-id="046f8-780">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-780">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="046f8-781">String</span><span class="sxs-lookup"><span data-stu-id="046f8-781">String</span></span>|<span data-ttu-id="046f8-782">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="046f8-782">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="046f8-783">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-783">Requirements</span></span>

|<span data-ttu-id="046f8-784">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-784">Requirement</span></span>| <span data-ttu-id="046f8-785">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-786">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-787">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-787">1.0</span></span>|
|[<span data-ttu-id="046f8-788">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-789">ReadItem</span></span>|
|[<span data-ttu-id="046f8-790">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-791">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-791">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="046f8-792">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="046f8-792">Returns:</span></span>

<span data-ttu-id="046f8-p152">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="046f8-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="046f8-795">Type : Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="046f8-795">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="046f8-796">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="046f8-796">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="046f8-797">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="046f8-797">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-798">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-798">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="046f8-p153">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="046f8-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="046f8-802">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="046f8-802">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="046f8-803">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="046f8-803">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="046f8-p154">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="046f8-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="046f8-806">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-806">Requirements</span></span>

|<span data-ttu-id="046f8-807">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-807">Requirement</span></span>| <span data-ttu-id="046f8-808">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-808">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-809">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-809">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-810">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-810">1.0</span></span>|
|[<span data-ttu-id="046f8-811">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-811">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-812">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-812">ReadItem</span></span>|
|[<span data-ttu-id="046f8-813">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-813">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-814">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-814">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="046f8-815">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="046f8-815">Returns:</span></span>

<span data-ttu-id="046f8-p155">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="046f8-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="046f8-818">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="046f8-818">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="046f8-819">Object</span><span class="sxs-lookup"><span data-stu-id="046f8-819">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="046f8-820">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-820">Example</span></span>

<span data-ttu-id="046f8-821">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="046f8-821">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="046f8-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="046f8-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="046f8-823">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="046f8-823">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="046f8-824">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="046f8-824">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="046f8-825">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="046f8-825">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="046f8-p156">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="046f8-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-828">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-828">Parameters:</span></span>

|<span data-ttu-id="046f8-829">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-829">Name</span></span>| <span data-ttu-id="046f8-830">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-830">Type</span></span>| <span data-ttu-id="046f8-831">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-831">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="046f8-832">String</span><span class="sxs-lookup"><span data-stu-id="046f8-832">String</span></span>|<span data-ttu-id="046f8-833">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="046f8-833">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="046f8-834">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-834">Requirements</span></span>

|<span data-ttu-id="046f8-835">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-835">Requirement</span></span>| <span data-ttu-id="046f8-836">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-836">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-837">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-837">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-838">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-838">1.0</span></span>|
|[<span data-ttu-id="046f8-839">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-839">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-840">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-840">ReadItem</span></span>|
|[<span data-ttu-id="046f8-841">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-841">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-842">Lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-842">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="046f8-843">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="046f8-843">Returns:</span></span>

<span data-ttu-id="046f8-844">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="046f8-844">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="046f8-845">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="046f8-845">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="046f8-846">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="046f8-846">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="046f8-847">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-847">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="046f8-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="046f8-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="046f8-849">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="046f8-849">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="046f8-p157">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="046f8-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-852">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-852">Parameters:</span></span>

|<span data-ttu-id="046f8-853">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-853">Name</span></span>| <span data-ttu-id="046f8-854">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-854">Type</span></span>| <span data-ttu-id="046f8-855">Attributs</span><span class="sxs-lookup"><span data-stu-id="046f8-855">Attributes</span></span>| <span data-ttu-id="046f8-856">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-856">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="046f8-857">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="046f8-857">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="046f8-p158">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="046f8-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="046f8-861">Object</span><span class="sxs-lookup"><span data-stu-id="046f8-861">Object</span></span>| <span data-ttu-id="046f8-862">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-862">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-863">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="046f8-863">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="046f8-864">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-864">Object</span></span>| <span data-ttu-id="046f8-865">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-865">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-866">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-866">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="046f8-867">fonction</span><span class="sxs-lookup"><span data-stu-id="046f8-867">function</span></span>||<span data-ttu-id="046f8-868">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-868">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="046f8-869">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="046f8-869">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="046f8-870">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="046f8-870">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="046f8-871">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-871">Requirements</span></span>

|<span data-ttu-id="046f8-872">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-872">Requirement</span></span>| <span data-ttu-id="046f8-873">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-874">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-875">1.2</span><span class="sxs-lookup"><span data-stu-id="046f8-875">1.2</span></span>|
|[<span data-ttu-id="046f8-876">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-876">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-877">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="046f8-877">ReadWriteItem</span></span>|
|[<span data-ttu-id="046f8-878">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-878">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-879">Composition</span><span class="sxs-lookup"><span data-stu-id="046f8-879">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="046f8-880">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="046f8-880">Returns:</span></span>

<span data-ttu-id="046f8-881">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="046f8-881">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="046f8-882">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="046f8-882">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="046f8-883">Chaîne</span><span class="sxs-lookup"><span data-stu-id="046f8-883">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="046f8-884">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-884">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="046f8-885">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="046f8-885">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="046f8-886">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="046f8-886">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="046f8-p160">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="046f8-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-890">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-890">Parameters:</span></span>

|<span data-ttu-id="046f8-891">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-891">Name</span></span>| <span data-ttu-id="046f8-892">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-892">Type</span></span>| <span data-ttu-id="046f8-893">Attributs</span><span class="sxs-lookup"><span data-stu-id="046f8-893">Attributes</span></span>| <span data-ttu-id="046f8-894">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-894">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="046f8-895">function</span><span class="sxs-lookup"><span data-stu-id="046f8-895">function</span></span>||<span data-ttu-id="046f8-896">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-896">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="046f8-897">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="046f8-897">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="046f8-898">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="046f8-898">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="046f8-899">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-899">Object</span></span>| <span data-ttu-id="046f8-900">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-900">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-901">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-901">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="046f8-902">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-902">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="046f8-903">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-903">Requirements</span></span>

|<span data-ttu-id="046f8-904">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-904">Requirement</span></span>| <span data-ttu-id="046f8-905">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-906">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-907">1.0</span><span class="sxs-lookup"><span data-stu-id="046f8-907">1.0</span></span>|
|[<span data-ttu-id="046f8-908">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="046f8-909">ReadItem</span></span>|
|[<span data-ttu-id="046f8-910">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-911">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="046f8-911">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-912">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-912">Example</span></span>

<span data-ttu-id="046f8-p163">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="046f8-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="046f8-916">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="046f8-916">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="046f8-917">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="046f8-917">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="046f8-p164">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="046f8-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-922">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-922">Parameters:</span></span>

|<span data-ttu-id="046f8-923">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-923">Name</span></span>| <span data-ttu-id="046f8-924">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-924">Type</span></span>| <span data-ttu-id="046f8-925">Attributs</span><span class="sxs-lookup"><span data-stu-id="046f8-925">Attributes</span></span>| <span data-ttu-id="046f8-926">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-926">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="046f8-927">String</span><span class="sxs-lookup"><span data-stu-id="046f8-927">String</span></span>||<span data-ttu-id="046f8-928">Identificateur de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="046f8-928">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="046f8-929">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-929">Object</span></span>| <span data-ttu-id="046f8-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-930">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-931">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="046f8-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="046f8-932">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-932">Object</span></span>| <span data-ttu-id="046f8-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-933">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-934">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="046f8-935">fonction</span><span class="sxs-lookup"><span data-stu-id="046f8-935">function</span></span>| <span data-ttu-id="046f8-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-936">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-937">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="046f8-938">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="046f8-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="046f8-939">Erreurs</span><span class="sxs-lookup"><span data-stu-id="046f8-939">Errors</span></span>

| <span data-ttu-id="046f8-940">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="046f8-940">Error code</span></span> | <span data-ttu-id="046f8-941">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="046f8-942">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="046f8-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="046f8-943">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-943">Requirements</span></span>

|<span data-ttu-id="046f8-944">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-944">Requirement</span></span>| <span data-ttu-id="046f8-945">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-946">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-947">1.1</span><span class="sxs-lookup"><span data-stu-id="046f8-947">1.1</span></span>|
|[<span data-ttu-id="046f8-948">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="046f8-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="046f8-950">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-951">Composition</span><span class="sxs-lookup"><span data-stu-id="046f8-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-952">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-952">Example</span></span>

<span data-ttu-id="046f8-953">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="046f8-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="046f8-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="046f8-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="046f8-955">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="046f8-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="046f8-p165">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="046f8-p165">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="046f8-959">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="046f8-959">Parameters:</span></span>

|<span data-ttu-id="046f8-960">Nom</span><span class="sxs-lookup"><span data-stu-id="046f8-960">Name</span></span>| <span data-ttu-id="046f8-961">Type</span><span class="sxs-lookup"><span data-stu-id="046f8-961">Type</span></span>| <span data-ttu-id="046f8-962">Attributs</span><span class="sxs-lookup"><span data-stu-id="046f8-962">Attributes</span></span>| <span data-ttu-id="046f8-963">Description</span><span class="sxs-lookup"><span data-stu-id="046f8-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="046f8-964">String</span><span class="sxs-lookup"><span data-stu-id="046f8-964">String</span></span>||<span data-ttu-id="046f8-p166">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="046f8-p166">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="046f8-968">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-968">Object</span></span>| <span data-ttu-id="046f8-969">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-969">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-970">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="046f8-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="046f8-971">Objet</span><span class="sxs-lookup"><span data-stu-id="046f8-971">Object</span></span>| <span data-ttu-id="046f8-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-972">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-973">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="046f8-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="046f8-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="046f8-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="046f8-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="046f8-975">&lt;optional&gt;</span></span>|<span data-ttu-id="046f8-p167">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="046f8-p167">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="046f8-p168">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="046f8-p168">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="046f8-980">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="046f8-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="046f8-981">fonction</span><span class="sxs-lookup"><span data-stu-id="046f8-981">function</span></span>||<span data-ttu-id="046f8-982">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="046f8-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="046f8-983">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="046f8-983">Requirements</span></span>

|<span data-ttu-id="046f8-984">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="046f8-984">Requirement</span></span>| <span data-ttu-id="046f8-985">Valeur</span><span class="sxs-lookup"><span data-stu-id="046f8-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="046f8-986">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="046f8-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="046f8-987">1.2</span><span class="sxs-lookup"><span data-stu-id="046f8-987">1.2</span></span>|
|[<span data-ttu-id="046f8-988">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="046f8-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="046f8-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="046f8-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="046f8-990">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="046f8-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="046f8-991">Composition</span><span class="sxs-lookup"><span data-stu-id="046f8-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="046f8-992">Exemple</span><span class="sxs-lookup"><span data-stu-id="046f8-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
