---
title: Office. Context. Mailbox-ensemble de conditions requises 1,2
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 542e8c9899c2d4a3c5b4546c3d5a73ba0d3c3a7e
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626998"
---
# <a name="mailbox"></a><span data-ttu-id="4d74e-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4d74e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4d74e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4d74e-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d74e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d74e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-105">Requirements</span></span>

|<span data-ttu-id="4d74e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-106">Requirement</span></span>| <span data-ttu-id="4d74e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-109">1.0</span></span>|
|[<span data-ttu-id="4d74e-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4d74e-111">Restricted</span></span>|
|[<span data-ttu-id="4d74e-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4d74e-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4d74e-114">Members and methods</span></span>

| <span data-ttu-id="4d74e-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4d74e-115">Member</span></span> | <span data-ttu-id="4d74e-116">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4d74e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="4d74e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="4d74e-118">Membre</span><span class="sxs-lookup"><span data-stu-id="4d74e-118">Member</span></span> |
| [<span data-ttu-id="4d74e-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4d74e-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="4d74e-120">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-120">Method</span></span> |
| [<span data-ttu-id="4d74e-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="4d74e-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="4d74e-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-122">Method</span></span> |
| [<span data-ttu-id="4d74e-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4d74e-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="4d74e-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-124">Method</span></span> |
| [<span data-ttu-id="4d74e-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="4d74e-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="4d74e-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-126">Method</span></span> |
| [<span data-ttu-id="4d74e-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4d74e-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="4d74e-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-128">Method</span></span> |
| [<span data-ttu-id="4d74e-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4d74e-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="4d74e-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-130">Method</span></span> |
| [<span data-ttu-id="4d74e-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4d74e-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="4d74e-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-132">Method</span></span> |
| [<span data-ttu-id="4d74e-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="4d74e-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="4d74e-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="4d74e-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4d74e-135">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4d74e-135">Namespaces</span></span>

<span data-ttu-id="4d74e-136">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d74e-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4d74e-137">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d74e-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4d74e-138">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4d74e-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4d74e-139">Members</span><span class="sxs-lookup"><span data-stu-id="4d74e-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4d74e-140">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="4d74e-140">ewsUrl: String</span></span>

<span data-ttu-id="4d74e-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4d74e-143">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d74e-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4d74e-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4d74e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="4d74e-146">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-146">Type</span></span>

*   <span data-ttu-id="4d74e-147">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d74e-148">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-148">Requirements</span></span>

|<span data-ttu-id="4d74e-149">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-149">Requirement</span></span>| <span data-ttu-id="4d74e-150">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-151">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-152">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-152">1.0</span></span>|
|[<span data-ttu-id="4d74e-153">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-154">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-155">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4d74e-157">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4d74e-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="4d74e-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="4d74e-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="4d74e-159">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="4d74e-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4d74e-p103">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p103">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4d74e-p104">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p104">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-165">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-165">Parameters</span></span>

|<span data-ttu-id="4d74e-166">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-166">Name</span></span>| <span data-ttu-id="4d74e-167">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-167">Type</span></span>| <span data-ttu-id="4d74e-168">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4d74e-169">Date</span><span class="sxs-lookup"><span data-stu-id="4d74e-169">Date</span></span>|<span data-ttu-id="4d74e-170">Objet Date</span><span class="sxs-lookup"><span data-stu-id="4d74e-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-171">Requirements</span></span>

|<span data-ttu-id="4d74e-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-172">Requirement</span></span>| <span data-ttu-id="4d74e-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-175">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-175">1.0</span></span>|
|[<span data-ttu-id="4d74e-176">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-177">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d74e-180">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d74e-180">Returns:</span></span>

<span data-ttu-id="4d74e-181">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="4d74e-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4d74e-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4d74e-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4d74e-183">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="4d74e-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4d74e-184">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="4d74e-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-185">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-185">Parameters</span></span>

|<span data-ttu-id="4d74e-186">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-186">Name</span></span>| <span data-ttu-id="4d74e-187">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-187">Type</span></span>| <span data-ttu-id="4d74e-188">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4d74e-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4d74e-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="4d74e-190">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="4d74e-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-191">Requirements</span></span>

|<span data-ttu-id="4d74e-192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-192">Requirement</span></span>| <span data-ttu-id="4d74e-193">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-195">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-195">1.0</span></span>|
|[<span data-ttu-id="4d74e-196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-197">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-199">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4d74e-200">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4d74e-200">Returns:</span></span>

<span data-ttu-id="4d74e-201">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="4d74e-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="4d74e-202">Type : Date</span><span class="sxs-lookup"><span data-stu-id="4d74e-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="4d74e-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-203">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="4d74e-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4d74e-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4d74e-205">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4d74e-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4d74e-206">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d74e-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4d74e-207">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4d74e-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4d74e-p105">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p105">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4d74e-210">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4d74e-211">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4d74e-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-212">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-212">Parameters</span></span>

|<span data-ttu-id="4d74e-213">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-213">Name</span></span>| <span data-ttu-id="4d74e-214">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-214">Type</span></span>| <span data-ttu-id="4d74e-215">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4d74e-216">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-216">String</span></span>|<span data-ttu-id="4d74e-217">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4d74e-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-218">Requirements</span></span>

|<span data-ttu-id="4d74e-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-219">Requirement</span></span>| <span data-ttu-id="4d74e-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-222">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-222">1.0</span></span>|
|[<span data-ttu-id="4d74e-223">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-224">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d74e-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="4d74e-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4d74e-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4d74e-229">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="4d74e-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4d74e-230">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d74e-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4d74e-231">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4d74e-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4d74e-232">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4d74e-233">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4d74e-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4d74e-p106">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-236">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-236">Parameters</span></span>

|<span data-ttu-id="4d74e-237">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-237">Name</span></span>| <span data-ttu-id="4d74e-238">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-238">Type</span></span>| <span data-ttu-id="4d74e-239">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4d74e-240">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-240">String</span></span>|<span data-ttu-id="4d74e-241">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="4d74e-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-242">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-242">Requirements</span></span>

|<span data-ttu-id="4d74e-243">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-243">Requirement</span></span>| <span data-ttu-id="4d74e-244">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-245">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-246">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-246">1.0</span></span>|
|[<span data-ttu-id="4d74e-247">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-248">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-249">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-250">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d74e-251">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4d74e-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4d74e-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4d74e-253">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="4d74e-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4d74e-254">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4d74e-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4d74e-p107">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4d74e-p108">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p108">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4d74e-p109">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4d74e-262">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="4d74e-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-263">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-263">Parameters</span></span>

|<span data-ttu-id="4d74e-264">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-264">Name</span></span>| <span data-ttu-id="4d74e-265">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-265">Type</span></span>| <span data-ttu-id="4d74e-266">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4d74e-267">Object</span><span class="sxs-lookup"><span data-stu-id="4d74e-267">Object</span></span> | <span data-ttu-id="4d74e-268">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d74e-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4d74e-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="4d74e-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="4d74e-p110">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4d74e-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="4d74e-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="4d74e-p111">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4d74e-275">Date</span><span class="sxs-lookup"><span data-stu-id="4d74e-275">Date</span></span> | <span data-ttu-id="4d74e-276">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d74e-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4d74e-277">Date</span><span class="sxs-lookup"><span data-stu-id="4d74e-277">Date</span></span> | <span data-ttu-id="4d74e-278">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4d74e-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4d74e-279">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-279">String</span></span> | <span data-ttu-id="4d74e-p112">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4d74e-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4d74e-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4d74e-p113">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4d74e-285">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-285">String</span></span> | <span data-ttu-id="4d74e-p114">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4d74e-288">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-288">String</span></span> | <span data-ttu-id="4d74e-p115">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4d74e-291">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-291">Requirements</span></span>

|<span data-ttu-id="4d74e-292">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-292">Requirement</span></span>| <span data-ttu-id="4d74e-293">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-294">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-295">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-295">1.0</span></span>|
|[<span data-ttu-id="4d74e-296">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-297">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-298">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-299">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d74e-300">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-300">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4d74e-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4d74e-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4d74e-302">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="4d74e-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4d74e-p116">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4d74e-305">Vous pouvez transmettre le jeton et soit un identificateur de pièce jointe, soit un identificateur d’élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="4d74e-305">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4d74e-306">Le système tiers utilise le jeton comme jeton d’autorisation du support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) de services Web Exchange (EWS) ou de [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour renvoyer une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="4d74e-306">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="4d74e-307">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4d74e-307">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4d74e-308">L’appel `getCallbackTokenAsync` de la méthode nécessite un niveau d’autorisation minimum de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="4d74e-308">Calling the `getCallbackTokenAsync` method requires a minimum permission level of **ReadItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-309">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-309">Parameters</span></span>

|<span data-ttu-id="4d74e-310">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-310">Name</span></span>| <span data-ttu-id="4d74e-311">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-311">Type</span></span>| <span data-ttu-id="4d74e-312">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d74e-312">Attributes</span></span>| <span data-ttu-id="4d74e-313">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4d74e-314">function</span><span class="sxs-lookup"><span data-stu-id="4d74e-314">function</span></span>||<span data-ttu-id="4d74e-315">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d74e-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4d74e-316">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d74e-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4d74e-317">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4d74e-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4d74e-318">Objet</span><span class="sxs-lookup"><span data-stu-id="4d74e-318">Object</span></span>| <span data-ttu-id="4d74e-319">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d74e-319">&lt;optional&gt;</span></span>|<span data-ttu-id="4d74e-320">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4d74e-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4d74e-321">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4d74e-321">Errors</span></span>

|<span data-ttu-id="4d74e-322">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4d74e-322">Error code</span></span>|<span data-ttu-id="4d74e-323">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4d74e-324">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4d74e-324">The request has failed.</span></span> <span data-ttu-id="4d74e-325">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4d74e-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4d74e-326">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4d74e-326">The Exchange server returned an error.</span></span> <span data-ttu-id="4d74e-327">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4d74e-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4d74e-328">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4d74e-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="4d74e-329">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="4d74e-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-330">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-330">Requirements</span></span>

|<span data-ttu-id="4d74e-331">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-331">Requirement</span></span>| <span data-ttu-id="4d74e-332">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-333">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-334">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-334">1.0</span></span>|
|[<span data-ttu-id="4d74e-335">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-336">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-337">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-338">Lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d74e-339">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-339">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4d74e-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4d74e-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4d74e-341">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="4d74e-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4d74e-342">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4d74e-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-343">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-343">Parameters</span></span>

|<span data-ttu-id="4d74e-344">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-344">Name</span></span>| <span data-ttu-id="4d74e-345">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-345">Type</span></span>| <span data-ttu-id="4d74e-346">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d74e-346">Attributes</span></span>| <span data-ttu-id="4d74e-347">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4d74e-348">function</span><span class="sxs-lookup"><span data-stu-id="4d74e-348">function</span></span>||<span data-ttu-id="4d74e-349">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d74e-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4d74e-350">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d74e-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4d74e-351">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4d74e-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4d74e-352">Objet</span><span class="sxs-lookup"><span data-stu-id="4d74e-352">Object</span></span>| <span data-ttu-id="4d74e-353">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d74e-353">&lt;optional&gt;</span></span>|<span data-ttu-id="4d74e-354">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4d74e-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4d74e-355">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4d74e-355">Errors</span></span>

|<span data-ttu-id="4d74e-356">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4d74e-356">Error code</span></span>|<span data-ttu-id="4d74e-357">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4d74e-358">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4d74e-358">The request has failed.</span></span> <span data-ttu-id="4d74e-359">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4d74e-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4d74e-360">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4d74e-360">The Exchange server returned an error.</span></span> <span data-ttu-id="4d74e-361">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4d74e-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4d74e-362">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4d74e-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="4d74e-363">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="4d74e-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-364">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-364">Requirements</span></span>

|<span data-ttu-id="4d74e-365">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-365">Requirement</span></span>| <span data-ttu-id="4d74e-366">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-367">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-368">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-368">1.0</span></span>|
|[<span data-ttu-id="4d74e-369">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4d74e-370">ReadItem</span></span>|
|[<span data-ttu-id="4d74e-371">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-372">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d74e-373">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-373">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4d74e-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4d74e-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4d74e-375">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d74e-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4d74e-376">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="4d74e-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4d74e-377">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="4d74e-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="4d74e-378">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="4d74e-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4d74e-379">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d74e-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4d74e-380">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="4d74e-380">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4d74e-381">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="4d74e-381">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4d74e-382">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="4d74e-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4d74e-383">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4d74e-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4d74e-p125">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4d74e-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4d74e-386">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="4d74e-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4d74e-387">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="4d74e-387">Version differences</span></span>

<span data-ttu-id="4d74e-388">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4d74e-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4d74e-p126">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4d74e-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4d74e-392">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d74e-392">Parameters</span></span>

|<span data-ttu-id="4d74e-393">Nom</span><span class="sxs-lookup"><span data-stu-id="4d74e-393">Name</span></span>| <span data-ttu-id="4d74e-394">Type</span><span class="sxs-lookup"><span data-stu-id="4d74e-394">Type</span></span>| <span data-ttu-id="4d74e-395">Attributs</span><span class="sxs-lookup"><span data-stu-id="4d74e-395">Attributes</span></span>| <span data-ttu-id="4d74e-396">Description</span><span class="sxs-lookup"><span data-stu-id="4d74e-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4d74e-397">String</span><span class="sxs-lookup"><span data-stu-id="4d74e-397">String</span></span>||<span data-ttu-id="4d74e-398">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="4d74e-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4d74e-399">function</span><span class="sxs-lookup"><span data-stu-id="4d74e-399">function</span></span>||<span data-ttu-id="4d74e-400">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4d74e-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4d74e-401">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4d74e-401">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4d74e-402">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="4d74e-402">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4d74e-403">Objet</span><span class="sxs-lookup"><span data-stu-id="4d74e-403">Object</span></span>| <span data-ttu-id="4d74e-404">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4d74e-404">&lt;optional&gt;</span></span>|<span data-ttu-id="4d74e-405">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4d74e-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d74e-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4d74e-406">Requirements</span></span>

|<span data-ttu-id="4d74e-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d74e-407">Requirement</span></span>| <span data-ttu-id="4d74e-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="4d74e-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d74e-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4d74e-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4d74e-410">1.0</span><span class="sxs-lookup"><span data-stu-id="4d74e-410">1.0</span></span>|
|[<span data-ttu-id="4d74e-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4d74e-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4d74e-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4d74e-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4d74e-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4d74e-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4d74e-414">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4d74e-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4d74e-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="4d74e-415">Example</span></span>

<span data-ttu-id="4d74e-416">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4d74e-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
