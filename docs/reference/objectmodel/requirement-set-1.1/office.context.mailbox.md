---
title: Office. Context. Mailbox-ensemble de conditions requises 1,1
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: d079ea8529255a8766fb3fd47b669dbb23d2ea64
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627061"
---
# <a name="mailbox"></a><span data-ttu-id="b803f-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="b803f-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="b803f-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="b803f-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="b803f-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b803f-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-105">Requirements</span></span>

|<span data-ttu-id="b803f-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-106">Requirement</span></span>| <span data-ttu-id="b803f-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-109">1.0</span></span>|
|[<span data-ttu-id="b803f-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b803f-111">Restricted</span></span>|
|[<span data-ttu-id="b803f-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b803f-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="b803f-114">Members and methods</span></span>

| <span data-ttu-id="b803f-115">Membre</span><span class="sxs-lookup"><span data-stu-id="b803f-115">Member</span></span> | <span data-ttu-id="b803f-116">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b803f-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="b803f-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="b803f-118">Membre</span><span class="sxs-lookup"><span data-stu-id="b803f-118">Member</span></span> |
| [<span data-ttu-id="b803f-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b803f-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="b803f-120">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-120">Method</span></span> |
| [<span data-ttu-id="b803f-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="b803f-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="b803f-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-122">Method</span></span> |
| [<span data-ttu-id="b803f-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b803f-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="b803f-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-124">Method</span></span> |
| [<span data-ttu-id="b803f-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="b803f-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="b803f-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-126">Method</span></span> |
| [<span data-ttu-id="b803f-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b803f-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="b803f-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-128">Method</span></span> |
| [<span data-ttu-id="b803f-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b803f-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="b803f-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-130">Method</span></span> |
| [<span data-ttu-id="b803f-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b803f-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="b803f-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-132">Method</span></span> |
| [<span data-ttu-id="b803f-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b803f-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="b803f-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="b803f-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b803f-135">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="b803f-135">Namespaces</span></span>

<span data-ttu-id="b803f-136">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b803f-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="b803f-137">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b803f-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="b803f-138">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b803f-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="b803f-139">Members</span><span class="sxs-lookup"><span data-stu-id="b803f-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="b803f-140">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="b803f-140">ewsUrl: String</span></span>

<span data-ttu-id="b803f-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="b803f-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b803f-143">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b803f-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b803f-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="b803f-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="b803f-146">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-146">Type</span></span>

*   <span data-ttu-id="b803f-147">String</span><span class="sxs-lookup"><span data-stu-id="b803f-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b803f-148">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-148">Requirements</span></span>

|<span data-ttu-id="b803f-149">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-149">Requirement</span></span>| <span data-ttu-id="b803f-150">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-151">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-152">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-152">1.0</span></span>|
|[<span data-ttu-id="b803f-153">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-154">ReadItem</span></span>|
|[<span data-ttu-id="b803f-155">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="b803f-157">Méthodes</span><span class="sxs-lookup"><span data-stu-id="b803f-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-11"></a><span data-ttu-id="b803f-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="b803f-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="b803f-159">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="b803f-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="b803f-p103">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b803f-p103">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="b803f-p104">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="b803f-p104">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-165">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-165">Parameters</span></span>

|<span data-ttu-id="b803f-166">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-166">Name</span></span>| <span data-ttu-id="b803f-167">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-167">Type</span></span>| <span data-ttu-id="b803f-168">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="b803f-169">Date</span><span class="sxs-lookup"><span data-stu-id="b803f-169">Date</span></span>|<span data-ttu-id="b803f-170">Objet Date</span><span class="sxs-lookup"><span data-stu-id="b803f-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-171">Requirements</span></span>

|<span data-ttu-id="b803f-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-172">Requirement</span></span>| <span data-ttu-id="b803f-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-175">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-175">1.0</span></span>|
|[<span data-ttu-id="b803f-176">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-177">ReadItem</span></span>|
|[<span data-ttu-id="b803f-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b803f-180">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b803f-180">Returns:</span></span>

<span data-ttu-id="b803f-181">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="b803f-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="b803f-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="b803f-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="b803f-183">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="b803f-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="b803f-184">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="b803f-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-185">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-185">Parameters</span></span>

|<span data-ttu-id="b803f-186">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-186">Name</span></span>| <span data-ttu-id="b803f-187">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-187">Type</span></span>| <span data-ttu-id="b803f-188">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="b803f-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b803f-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)|<span data-ttu-id="b803f-190">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="b803f-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-191">Requirements</span></span>

|<span data-ttu-id="b803f-192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-192">Requirement</span></span>| <span data-ttu-id="b803f-193">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-195">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-195">1.0</span></span>|
|[<span data-ttu-id="b803f-196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-197">ReadItem</span></span>|
|[<span data-ttu-id="b803f-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-199">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b803f-200">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="b803f-200">Returns:</span></span>

<span data-ttu-id="b803f-201">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="b803f-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="b803f-202">Type : Date</span><span class="sxs-lookup"><span data-stu-id="b803f-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="b803f-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-203">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="b803f-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b803f-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="b803f-205">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="b803f-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b803f-206">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b803f-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b803f-207">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="b803f-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b803f-p105">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="b803f-p105">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="b803f-210">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="b803f-211">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="b803f-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-212">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-212">Parameters</span></span>

|<span data-ttu-id="b803f-213">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-213">Name</span></span>| <span data-ttu-id="b803f-214">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-214">Type</span></span>| <span data-ttu-id="b803f-215">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b803f-216">String</span><span class="sxs-lookup"><span data-stu-id="b803f-216">String</span></span>|<span data-ttu-id="b803f-217">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="b803f-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-218">Requirements</span></span>

|<span data-ttu-id="b803f-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-219">Requirement</span></span>| <span data-ttu-id="b803f-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-222">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-222">1.0</span></span>|
|[<span data-ttu-id="b803f-223">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-224">ReadItem</span></span>|
|[<span data-ttu-id="b803f-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b803f-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="b803f-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b803f-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="b803f-229">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="b803f-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="b803f-230">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b803f-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b803f-231">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="b803f-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b803f-232">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="b803f-233">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="b803f-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="b803f-p106">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b803f-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-236">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-236">Parameters</span></span>

|<span data-ttu-id="b803f-237">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-237">Name</span></span>| <span data-ttu-id="b803f-238">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-238">Type</span></span>| <span data-ttu-id="b803f-239">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b803f-240">String</span><span class="sxs-lookup"><span data-stu-id="b803f-240">String</span></span>|<span data-ttu-id="b803f-241">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="b803f-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-242">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-242">Requirements</span></span>

|<span data-ttu-id="b803f-243">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-243">Requirement</span></span>| <span data-ttu-id="b803f-244">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-245">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-246">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-246">1.0</span></span>|
|[<span data-ttu-id="b803f-247">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-248">ReadItem</span></span>|
|[<span data-ttu-id="b803f-249">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-250">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b803f-251">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="b803f-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="b803f-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="b803f-253">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="b803f-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b803f-254">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="b803f-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b803f-p107">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="b803f-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="b803f-p108">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="b803f-p108">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="b803f-p109">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="b803f-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="b803f-262">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="b803f-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-263">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-263">Parameters</span></span>

|<span data-ttu-id="b803f-264">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-264">Name</span></span>| <span data-ttu-id="b803f-265">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-265">Type</span></span>| <span data-ttu-id="b803f-266">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="b803f-267">Object</span><span class="sxs-lookup"><span data-stu-id="b803f-267">Object</span></span> | <span data-ttu-id="b803f-268">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b803f-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="b803f-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span><span class="sxs-lookup"><span data-stu-id="b803f-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span></span> | <span data-ttu-id="b803f-p110">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="b803f-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span><span class="sxs-lookup"><span data-stu-id="b803f-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span></span> | <span data-ttu-id="b803f-p111">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="b803f-275">Date</span><span class="sxs-lookup"><span data-stu-id="b803f-275">Date</span></span> | <span data-ttu-id="b803f-276">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b803f-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="b803f-277">Date</span><span class="sxs-lookup"><span data-stu-id="b803f-277">Date</span></span> | <span data-ttu-id="b803f-278">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="b803f-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="b803f-279">String</span><span class="sxs-lookup"><span data-stu-id="b803f-279">String</span></span> | <span data-ttu-id="b803f-p112">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="b803f-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="b803f-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="b803f-p113">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="b803f-285">String</span><span class="sxs-lookup"><span data-stu-id="b803f-285">String</span></span> | <span data-ttu-id="b803f-p114">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="b803f-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="b803f-288">String</span><span class="sxs-lookup"><span data-stu-id="b803f-288">String</span></span> | <span data-ttu-id="b803f-p115">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="b803f-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b803f-291">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-291">Requirements</span></span>

|<span data-ttu-id="b803f-292">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-292">Requirement</span></span>| <span data-ttu-id="b803f-293">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-294">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-295">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-295">1.0</span></span>|
|[<span data-ttu-id="b803f-296">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-297">ReadItem</span></span>|
|[<span data-ttu-id="b803f-298">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-299">Lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b803f-300">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-300">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="b803f-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b803f-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b803f-302">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b803f-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="b803f-p116">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="b803f-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="b803f-305">Vous pouvez transmettre le jeton et soit un identificateur de pièce jointe, soit un identificateur d’élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="b803f-305">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="b803f-306">Le système tiers utilise le jeton comme jeton d’autorisation du support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) de services Web Exchange (EWS) ou de [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour renvoyer une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="b803f-306">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="b803f-307">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="b803f-307">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b803f-308">L’appel `getCallbackTokenAsync` de la méthode nécessite un niveau d’autorisation minimum de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="b803f-308">Calling the `getCallbackTokenAsync` method requires a minimum permission level of **ReadItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-309">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-309">Parameters</span></span>

|<span data-ttu-id="b803f-310">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-310">Name</span></span>| <span data-ttu-id="b803f-311">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-311">Type</span></span>| <span data-ttu-id="b803f-312">Attributs</span><span class="sxs-lookup"><span data-stu-id="b803f-312">Attributes</span></span>| <span data-ttu-id="b803f-313">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b803f-314">function</span><span class="sxs-lookup"><span data-stu-id="b803f-314">function</span></span>||<span data-ttu-id="b803f-315">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b803f-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b803f-316">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b803f-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b803f-317">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="b803f-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="b803f-318">Objet</span><span class="sxs-lookup"><span data-stu-id="b803f-318">Object</span></span>| <span data-ttu-id="b803f-319">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b803f-319">&lt;optional&gt;</span></span>|<span data-ttu-id="b803f-320">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b803f-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b803f-321">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b803f-321">Errors</span></span>

|<span data-ttu-id="b803f-322">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b803f-322">Error code</span></span>|<span data-ttu-id="b803f-323">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b803f-324">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="b803f-324">The request has failed.</span></span> <span data-ttu-id="b803f-325">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="b803f-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b803f-326">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="b803f-326">The Exchange server returned an error.</span></span> <span data-ttu-id="b803f-327">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="b803f-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b803f-328">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="b803f-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="b803f-329">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="b803f-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-330">Requirements</span></span>

|<span data-ttu-id="b803f-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-331">Requirement</span></span>| <span data-ttu-id="b803f-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-334">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-334">1.0</span></span>|
|[<span data-ttu-id="b803f-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-336">ReadItem</span></span>|
|[<span data-ttu-id="b803f-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-338">Lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b803f-339">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-339">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="b803f-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b803f-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b803f-341">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="b803f-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="b803f-342">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="b803f-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-343">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-343">Parameters</span></span>

|<span data-ttu-id="b803f-344">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-344">Name</span></span>| <span data-ttu-id="b803f-345">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-345">Type</span></span>| <span data-ttu-id="b803f-346">Attributs</span><span class="sxs-lookup"><span data-stu-id="b803f-346">Attributes</span></span>| <span data-ttu-id="b803f-347">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b803f-348">function</span><span class="sxs-lookup"><span data-stu-id="b803f-348">function</span></span>||<span data-ttu-id="b803f-349">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b803f-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b803f-350">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b803f-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b803f-351">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="b803f-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="b803f-352">Objet</span><span class="sxs-lookup"><span data-stu-id="b803f-352">Object</span></span>| <span data-ttu-id="b803f-353">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b803f-353">&lt;optional&gt;</span></span>|<span data-ttu-id="b803f-354">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b803f-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b803f-355">Erreurs</span><span class="sxs-lookup"><span data-stu-id="b803f-355">Errors</span></span>

|<span data-ttu-id="b803f-356">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="b803f-356">Error code</span></span>|<span data-ttu-id="b803f-357">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b803f-358">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="b803f-358">The request has failed.</span></span> <span data-ttu-id="b803f-359">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="b803f-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b803f-360">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="b803f-360">The Exchange server returned an error.</span></span> <span data-ttu-id="b803f-361">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="b803f-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b803f-362">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="b803f-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="b803f-363">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="b803f-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-364">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-364">Requirements</span></span>

|<span data-ttu-id="b803f-365">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-365">Requirement</span></span>| <span data-ttu-id="b803f-366">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-367">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-368">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-368">1.0</span></span>|
|[<span data-ttu-id="b803f-369">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b803f-370">ReadItem</span></span>|
|[<span data-ttu-id="b803f-371">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-372">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b803f-373">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-373">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="b803f-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b803f-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="b803f-375">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b803f-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="b803f-376">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="b803f-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="b803f-377">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="b803f-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="b803f-378">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="b803f-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="b803f-379">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b803f-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="b803f-380">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="b803f-380">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="b803f-381">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="b803f-381">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="b803f-382">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="b803f-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="b803f-383">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="b803f-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="b803f-p125">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="b803f-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="b803f-386">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="b803f-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="b803f-387">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="b803f-387">Version differences</span></span>

<span data-ttu-id="b803f-388">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="b803f-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="b803f-p126">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="b803f-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b803f-392">Paramètres</span><span class="sxs-lookup"><span data-stu-id="b803f-392">Parameters</span></span>

|<span data-ttu-id="b803f-393">Nom</span><span class="sxs-lookup"><span data-stu-id="b803f-393">Name</span></span>| <span data-ttu-id="b803f-394">Type</span><span class="sxs-lookup"><span data-stu-id="b803f-394">Type</span></span>| <span data-ttu-id="b803f-395">Attributs</span><span class="sxs-lookup"><span data-stu-id="b803f-395">Attributes</span></span>| <span data-ttu-id="b803f-396">Description</span><span class="sxs-lookup"><span data-stu-id="b803f-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b803f-397">String</span><span class="sxs-lookup"><span data-stu-id="b803f-397">String</span></span>||<span data-ttu-id="b803f-398">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="b803f-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="b803f-399">function</span><span class="sxs-lookup"><span data-stu-id="b803f-399">function</span></span>||<span data-ttu-id="b803f-400">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b803f-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b803f-401">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b803f-401">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="b803f-402">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="b803f-402">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="b803f-403">Objet</span><span class="sxs-lookup"><span data-stu-id="b803f-403">Object</span></span>| <span data-ttu-id="b803f-404">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b803f-404">&lt;optional&gt;</span></span>|<span data-ttu-id="b803f-405">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b803f-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b803f-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b803f-406">Requirements</span></span>

|<span data-ttu-id="b803f-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b803f-407">Requirement</span></span>| <span data-ttu-id="b803f-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="b803f-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="b803f-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b803f-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b803f-410">1.0</span><span class="sxs-lookup"><span data-stu-id="b803f-410">1.0</span></span>|
|[<span data-ttu-id="b803f-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b803f-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b803f-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="b803f-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="b803f-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b803f-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b803f-414">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b803f-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b803f-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="b803f-415">Example</span></span>

<span data-ttu-id="b803f-416">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="b803f-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
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
