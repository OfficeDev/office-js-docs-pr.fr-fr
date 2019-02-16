---
title: Office.context – ensemble de conditions requises 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 5d3cc3fbcc233e8afd0d0a6f3a3ae80534eb3c28
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068216"
---
# <a name="mailbox"></a><span data-ttu-id="46f25-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="46f25-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="46f25-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="46f25-104">Permet d’accéder au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="46f25-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="46f25-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-105">Requirements</span></span>

|<span data-ttu-id="46f25-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-106">Requirement</span></span>| <span data-ttu-id="46f25-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-109">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-109">1.0</span></span>|
|[<span data-ttu-id="46f25-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="46f25-111">Restricted</span></span>|
|[<span data-ttu-id="46f25-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="46f25-114">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="46f25-114">Namespaces</span></span>

<span data-ttu-id="46f25-115">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="46f25-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="46f25-116">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="46f25-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="46f25-117">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="46f25-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="46f25-118">Membres</span><span class="sxs-lookup"><span data-stu-id="46f25-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="46f25-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="46f25-119">ewsUrl :String</span></span>

<span data-ttu-id="46f25-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="46f25-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="46f25-122">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="46f25-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="46f25-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="46f25-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="46f25-125">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-125">Type</span></span>

*   <span data-ttu-id="46f25-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="46f25-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="46f25-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-127">Requirements</span></span>

|<span data-ttu-id="46f25-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-128">Requirement</span></span>| <span data-ttu-id="46f25-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-131">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-131">1.0</span></span>|
|[<span data-ttu-id="46f25-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-133">ReadItem</span></span>|
|[<span data-ttu-id="46f25-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-135">Lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="46f25-136">Méthodes</span><span class="sxs-lookup"><span data-stu-id="46f25-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime"></a><span data-ttu-id="46f25-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="46f25-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span></span>

<span data-ttu-id="46f25-138">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="46f25-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="46f25-p103">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="46f25-p103">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="46f25-p104">Si l’application de messagerie est en cours d’exécution dans Outlook, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook Web App, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="46f25-p104">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-144">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-144">Parameters</span></span>

|<span data-ttu-id="46f25-145">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-145">Name</span></span>| <span data-ttu-id="46f25-146">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-146">Type</span></span>| <span data-ttu-id="46f25-147">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="46f25-148">Date</span><span class="sxs-lookup"><span data-stu-id="46f25-148">Date</span></span>|<span data-ttu-id="46f25-149">Objet Date</span><span class="sxs-lookup"><span data-stu-id="46f25-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-150">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-150">Requirements</span></span>

|<span data-ttu-id="46f25-151">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-151">Requirement</span></span>| <span data-ttu-id="46f25-152">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-153">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-154">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-154">1.0</span></span>|
|[<span data-ttu-id="46f25-155">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-156">ReadItem</span></span>|
|[<span data-ttu-id="46f25-157">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-158">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-158">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="46f25-159">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="46f25-159">Returns:</span></span>

<span data-ttu-id="46f25-160">Type : [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="46f25-160">Type: [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="46f25-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="46f25-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="46f25-162">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="46f25-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="46f25-163">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="46f25-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-164">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-164">Parameters</span></span>

|<span data-ttu-id="46f25-165">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-165">Name</span></span>| <span data-ttu-id="46f25-166">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-166">Type</span></span>| <span data-ttu-id="46f25-167">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="46f25-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="46f25-168">LocalClientTime</span></span>](/javascript/api/outlook_1_2/office.LocalClientTime)|<span data-ttu-id="46f25-169">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="46f25-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-170">Requirements</span></span>

|<span data-ttu-id="46f25-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-171">Requirement</span></span>| <span data-ttu-id="46f25-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-174">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-174">1.0</span></span>|
|[<span data-ttu-id="46f25-175">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-176">ReadItem</span></span>|
|[<span data-ttu-id="46f25-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-178">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-178">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="46f25-179">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="46f25-179">Returns:</span></span>

<span data-ttu-id="46f25-180">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="46f25-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="46f25-181">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="46f25-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="46f25-182">Date</span><span class="sxs-lookup"><span data-stu-id="46f25-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="46f25-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="46f25-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="46f25-184">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="46f25-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="46f25-185">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="46f25-185">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="46f25-186">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="46f25-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="46f25-p105">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="46f25-p105">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="46f25-189">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="46f25-190">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="46f25-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-191">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-191">Parameters</span></span>

|<span data-ttu-id="46f25-192">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-192">Name</span></span>| <span data-ttu-id="46f25-193">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-193">Type</span></span>| <span data-ttu-id="46f25-194">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="46f25-195">String</span><span class="sxs-lookup"><span data-stu-id="46f25-195">String</span></span>|<span data-ttu-id="46f25-196">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="46f25-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-197">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-197">Requirements</span></span>

|<span data-ttu-id="46f25-198">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-198">Requirement</span></span>| <span data-ttu-id="46f25-199">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-200">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-201">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-201">1.0</span></span>|
|[<span data-ttu-id="46f25-202">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-202">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-203">ReadItem</span></span>|
|[<span data-ttu-id="46f25-204">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-204">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-205">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-205">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="46f25-206">Exemple</span><span class="sxs-lookup"><span data-stu-id="46f25-206">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="46f25-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="46f25-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="46f25-208">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="46f25-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="46f25-209">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="46f25-209">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="46f25-210">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="46f25-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="46f25-211">Dans Outlook Web App, cette méthode ouvre le formulaire indiqué uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="46f25-212">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="46f25-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="46f25-p106">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="46f25-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-215">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-215">Parameters</span></span>

|<span data-ttu-id="46f25-216">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-216">Name</span></span>| <span data-ttu-id="46f25-217">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-217">Type</span></span>| <span data-ttu-id="46f25-218">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="46f25-219">String</span><span class="sxs-lookup"><span data-stu-id="46f25-219">String</span></span>|<span data-ttu-id="46f25-220">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="46f25-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-221">Requirements</span></span>

|<span data-ttu-id="46f25-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-222">Requirement</span></span>| <span data-ttu-id="46f25-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-225">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-225">1.0</span></span>|
|[<span data-ttu-id="46f25-226">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-227">ReadItem</span></span>|
|[<span data-ttu-id="46f25-228">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-229">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-229">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="46f25-230">Exemple</span><span class="sxs-lookup"><span data-stu-id="46f25-230">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="46f25-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="46f25-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="46f25-232">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="46f25-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="46f25-233">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="46f25-233">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="46f25-p107">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="46f25-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="46f25-p108">Dans Outlook Web App et OWA pour les périphériques, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="46f25-p108">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="46f25-p109">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="46f25-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="46f25-241">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="46f25-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-242">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-242">Parameters</span></span>

|<span data-ttu-id="46f25-243">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-243">Name</span></span>| <span data-ttu-id="46f25-244">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-244">Type</span></span>| <span data-ttu-id="46f25-245">object</span><span class="sxs-lookup"><span data-stu-id="46f25-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="46f25-246">Objet</span><span class="sxs-lookup"><span data-stu-id="46f25-246">Object</span></span> | <span data-ttu-id="46f25-247">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="46f25-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="46f25-248">Tableau.&lt;Chaîne&gt; &#124; Tableau.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="46f25-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="46f25-p110">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="46f25-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="46f25-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="46f25-p111">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="46f25-254">Date</span><span class="sxs-lookup"><span data-stu-id="46f25-254">Date</span></span> | <span data-ttu-id="46f25-255">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="46f25-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="46f25-256">Date</span><span class="sxs-lookup"><span data-stu-id="46f25-256">Date</span></span> | <span data-ttu-id="46f25-257">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="46f25-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="46f25-258">String</span><span class="sxs-lookup"><span data-stu-id="46f25-258">String</span></span> | <span data-ttu-id="46f25-p112">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="46f25-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="46f25-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="46f25-p113">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="46f25-264">String</span><span class="sxs-lookup"><span data-stu-id="46f25-264">String</span></span> | <span data-ttu-id="46f25-p114">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="46f25-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="46f25-267">String</span><span class="sxs-lookup"><span data-stu-id="46f25-267">String</span></span> | <span data-ttu-id="46f25-p115">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="46f25-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="46f25-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-270">Requirements</span></span>

|<span data-ttu-id="46f25-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-271">Requirement</span></span>| <span data-ttu-id="46f25-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-274">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-274">1.0</span></span>|
|[<span data-ttu-id="46f25-275">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-276">ReadItem</span></span>|
|[<span data-ttu-id="46f25-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-278">Lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="46f25-279">Exemple</span><span class="sxs-lookup"><span data-stu-id="46f25-279">Example</span></span>

```javascript
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="46f25-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="46f25-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="46f25-281">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="46f25-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="46f25-p116">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="46f25-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="46f25-p117">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="46f25-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="46f25-287">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="46f25-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-288">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-288">Parameters</span></span>

|<span data-ttu-id="46f25-289">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-289">Name</span></span>| <span data-ttu-id="46f25-290">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-290">Type</span></span>| <span data-ttu-id="46f25-291">Attributs</span><span class="sxs-lookup"><span data-stu-id="46f25-291">Attributes</span></span>| <span data-ttu-id="46f25-292">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="46f25-293">function</span><span class="sxs-lookup"><span data-stu-id="46f25-293">function</span></span>||<span data-ttu-id="46f25-294">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="46f25-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="46f25-295">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="46f25-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="46f25-296">Object</span><span class="sxs-lookup"><span data-stu-id="46f25-296">Object</span></span>| <span data-ttu-id="46f25-297">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="46f25-297">&lt;optional&gt;</span></span>|<span data-ttu-id="46f25-298">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="46f25-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-299">Requirements</span></span>

|<span data-ttu-id="46f25-300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-300">Requirement</span></span>| <span data-ttu-id="46f25-301">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-303">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-303">1.0</span></span>|
|[<span data-ttu-id="46f25-304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-305">ReadItem</span></span>|
|[<span data-ttu-id="46f25-306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-307">Lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="46f25-308">Exemple</span><span class="sxs-lookup"><span data-stu-id="46f25-308">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="46f25-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="46f25-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="46f25-310">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="46f25-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="46f25-311">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="46f25-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-312">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-312">Parameters</span></span>

|<span data-ttu-id="46f25-313">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-313">Name</span></span>| <span data-ttu-id="46f25-314">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-314">Type</span></span>| <span data-ttu-id="46f25-315">Attributs</span><span class="sxs-lookup"><span data-stu-id="46f25-315">Attributes</span></span>| <span data-ttu-id="46f25-316">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="46f25-317">function</span><span class="sxs-lookup"><span data-stu-id="46f25-317">function</span></span>||<span data-ttu-id="46f25-318">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="46f25-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="46f25-319">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="46f25-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="46f25-320">Object</span><span class="sxs-lookup"><span data-stu-id="46f25-320">Object</span></span>| <span data-ttu-id="46f25-321">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="46f25-321">&lt;optional&gt;</span></span>|<span data-ttu-id="46f25-322">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="46f25-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-323">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-323">Requirements</span></span>

|<span data-ttu-id="46f25-324">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-324">Requirement</span></span>| <span data-ttu-id="46f25-325">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-326">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-327">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-327">1.0</span></span>|
|[<span data-ttu-id="46f25-328">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46f25-329">ReadItem</span></span>|
|[<span data-ttu-id="46f25-330">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-331">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-331">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="46f25-332">Exemple</span><span class="sxs-lookup"><span data-stu-id="46f25-332">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="46f25-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="46f25-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="46f25-334">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="46f25-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="46f25-335">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="46f25-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="46f25-336">dans Outlook pour iOS ou Outlook pour Android ;</span><span class="sxs-lookup"><span data-stu-id="46f25-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="46f25-337">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="46f25-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="46f25-338">Dans ces cas de figure, les compléments doivent [utiliser les API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="46f25-338">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="46f25-339">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="46f25-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="46f25-340">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="46f25-340">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="46f25-341">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="46f25-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="46f25-342">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="46f25-342">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="46f25-p119">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="46f25-p119">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="46f25-345">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="46f25-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="46f25-346">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="46f25-346">Version differences</span></span>

<span data-ttu-id="46f25-347">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="46f25-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="46f25-p120">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="46f25-p120">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="46f25-351">Paramètres</span><span class="sxs-lookup"><span data-stu-id="46f25-351">Parameters</span></span>

|<span data-ttu-id="46f25-352">Nom</span><span class="sxs-lookup"><span data-stu-id="46f25-352">Name</span></span>| <span data-ttu-id="46f25-353">Type</span><span class="sxs-lookup"><span data-stu-id="46f25-353">Type</span></span>| <span data-ttu-id="46f25-354">Attributs</span><span class="sxs-lookup"><span data-stu-id="46f25-354">Attributes</span></span>| <span data-ttu-id="46f25-355">Description</span><span class="sxs-lookup"><span data-stu-id="46f25-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="46f25-356">String</span><span class="sxs-lookup"><span data-stu-id="46f25-356">String</span></span>||<span data-ttu-id="46f25-357">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="46f25-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="46f25-358">function</span><span class="sxs-lookup"><span data-stu-id="46f25-358">function</span></span>||<span data-ttu-id="46f25-359">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="46f25-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="46f25-360">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="46f25-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="46f25-361">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="46f25-361">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="46f25-362">Objet</span><span class="sxs-lookup"><span data-stu-id="46f25-362">Object</span></span>| <span data-ttu-id="46f25-363">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="46f25-363">&lt;optional&gt;</span></span>|<span data-ttu-id="46f25-364">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="46f25-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46f25-365">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46f25-365">Requirements</span></span>

|<span data-ttu-id="46f25-366">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46f25-366">Requirement</span></span>| <span data-ttu-id="46f25-367">Valeur</span><span class="sxs-lookup"><span data-stu-id="46f25-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="46f25-368">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46f25-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46f25-369">1.0</span><span class="sxs-lookup"><span data-stu-id="46f25-369">1.0</span></span>|
|[<span data-ttu-id="46f25-370">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46f25-370">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46f25-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="46f25-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="46f25-372">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46f25-372">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46f25-373">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46f25-373">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="46f25-374">Exemple</span><span class="sxs-lookup"><span data-stu-id="46f25-374">Example</span></span>

<span data-ttu-id="46f25-375">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="46f25-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```
