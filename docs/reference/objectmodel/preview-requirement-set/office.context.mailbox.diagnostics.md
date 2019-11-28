---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 492e292737417854adfaf98feb2b67788933d874
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629201"
---
# <a name="diagnostics"></a><span data-ttu-id="ee7a5-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="ee7a5-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="ee7a5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="ee7a5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="ee7a5-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee7a5-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ee7a5-105">Requirements</span></span>

|<span data-ttu-id="ee7a5-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ee7a5-106">Requirement</span></span>| <span data-ttu-id="ee7a5-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="ee7a5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee7a5-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ee7a5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee7a5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-109">1.0</span></span>|
|[<span data-ttu-id="ee7a5-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ee7a5-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee7a5-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-111">ReadItem</span></span>|
|[<span data-ttu-id="ee7a5-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ee7a5-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ee7a5-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ee7a5-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ee7a5-114">Properties</span></span>

| <span data-ttu-id="ee7a5-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="ee7a5-115">Property</span></span> | <span data-ttu-id="ee7a5-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="ee7a5-116">Minimum</span></span><br><span data-ttu-id="ee7a5-117">niveau d’autorisation</span><span class="sxs-lookup"><span data-stu-id="ee7a5-117">permission level</span></span> | <span data-ttu-id="ee7a5-118">Modes</span><span class="sxs-lookup"><span data-stu-id="ee7a5-118">Modes</span></span> | <span data-ttu-id="ee7a5-119">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ee7a5-119">Return type</span></span> | <span data-ttu-id="ee7a5-120">Minimale</span><span class="sxs-lookup"><span data-stu-id="ee7a5-120">Minimum</span></span><br><span data-ttu-id="ee7a5-121">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ee7a5-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="ee7a5-122">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="ee7a5-122">hostName</span></span>](#hostname-string) | <span data-ttu-id="ee7a5-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-123">ReadItem</span></span> | <span data-ttu-id="ee7a5-124">Composition</span><span class="sxs-lookup"><span data-stu-id="ee7a5-124">Compose</span></span><br><span data-ttu-id="ee7a5-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-125">Read</span></span> | <span data-ttu-id="ee7a5-126">String</span><span class="sxs-lookup"><span data-stu-id="ee7a5-126">String</span></span> | <span data-ttu-id="ee7a5-127">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-127">1.0</span></span> |
| [<span data-ttu-id="ee7a5-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="ee7a5-128">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="ee7a5-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-129">ReadItem</span></span> | <span data-ttu-id="ee7a5-130">Composition</span><span class="sxs-lookup"><span data-stu-id="ee7a5-130">Compose</span></span><br><span data-ttu-id="ee7a5-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-131">Read</span></span> | <span data-ttu-id="ee7a5-132">String</span><span class="sxs-lookup"><span data-stu-id="ee7a5-132">String</span></span> | <span data-ttu-id="ee7a5-133">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-133">1.0</span></span> |
| [<span data-ttu-id="ee7a5-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="ee7a5-134">OWAView</span></span>](#owaview-string) | <span data-ttu-id="ee7a5-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-135">ReadItem</span></span> | <span data-ttu-id="ee7a5-136">Composition</span><span class="sxs-lookup"><span data-stu-id="ee7a5-136">Compose</span></span><br><span data-ttu-id="ee7a5-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-137">Read</span></span> | <span data-ttu-id="ee7a5-138">String</span><span class="sxs-lookup"><span data-stu-id="ee7a5-138">String</span></span> | <span data-ttu-id="ee7a5-139">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-139">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="ee7a5-140">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="ee7a5-140">Property details</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="ee7a5-141">NomHôte : chaîne</span><span class="sxs-lookup"><span data-stu-id="ee7a5-141">hostName: String</span></span>

<span data-ttu-id="ee7a5-142">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-142">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="ee7a5-143">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-143">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="ee7a5-144">La `Outlook` valeur est renvoyée pour Outlook sur les clients de bureau (par exemple, Windows et Mac).</span><span class="sxs-lookup"><span data-stu-id="ee7a5-144">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="ee7a5-145">Type</span><span class="sxs-lookup"><span data-stu-id="ee7a5-145">Type</span></span>

*   <span data-ttu-id="ee7a5-146">String</span><span class="sxs-lookup"><span data-stu-id="ee7a5-146">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee7a5-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ee7a5-147">Requirements</span></span>

|<span data-ttu-id="ee7a5-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ee7a5-148">Requirement</span></span>| <span data-ttu-id="ee7a5-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="ee7a5-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee7a5-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ee7a5-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee7a5-151">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-151">1.0</span></span>|
|[<span data-ttu-id="ee7a5-152">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ee7a5-152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee7a5-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-153">ReadItem</span></span>|
|[<span data-ttu-id="ee7a5-154">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ee7a5-154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ee7a5-155">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-155">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="ee7a5-156">hostVersion : chaîne</span><span class="sxs-lookup"><span data-stu-id="ee7a5-156">hostVersion: String</span></span>

<span data-ttu-id="ee7a5-157">Obtient une valeur de type String qui représente la version de l’application hôte ou du serveur Exchange (par exemple, « 15.0.468.0 »).</span><span class="sxs-lookup"><span data-stu-id="ee7a5-157">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="ee7a5-158">Si le complément de messagerie est exécuté sur un ordinateur de bureau ou un client mobile Outlook `hostVersion` , la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-158">If the mail add-in is running on an Outlook desktop or mobile client, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="ee7a5-159">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-159">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="ee7a5-160">Type</span><span class="sxs-lookup"><span data-stu-id="ee7a5-160">Type</span></span>

*   <span data-ttu-id="ee7a5-161">String</span><span class="sxs-lookup"><span data-stu-id="ee7a5-161">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee7a5-162">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ee7a5-162">Requirements</span></span>

|<span data-ttu-id="ee7a5-163">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ee7a5-163">Requirement</span></span>| <span data-ttu-id="ee7a5-164">Valeur</span><span class="sxs-lookup"><span data-stu-id="ee7a5-164">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee7a5-165">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ee7a5-165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee7a5-166">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-166">1.0</span></span>|
|[<span data-ttu-id="ee7a5-167">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ee7a5-167">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee7a5-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-168">ReadItem</span></span>|
|[<span data-ttu-id="ee7a5-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ee7a5-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ee7a5-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="ee7a5-171">OWAView : chaîne</span><span class="sxs-lookup"><span data-stu-id="ee7a5-171">OWAView: String</span></span>

<span data-ttu-id="ee7a5-172">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-172">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="ee7a5-173">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-173">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="ee7a5-174">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-174">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="ee7a5-175">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="ee7a5-175">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="ee7a5-176">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-176">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="ee7a5-177">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-177">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="ee7a5-178">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-178">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="ee7a5-179">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-179">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="ee7a5-180">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-180">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="ee7a5-181">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="ee7a5-181">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="ee7a5-182">Type</span><span class="sxs-lookup"><span data-stu-id="ee7a5-182">Type</span></span>

*   <span data-ttu-id="ee7a5-183">String</span><span class="sxs-lookup"><span data-stu-id="ee7a5-183">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee7a5-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ee7a5-184">Requirements</span></span>

|<span data-ttu-id="ee7a5-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ee7a5-185">Requirement</span></span>| <span data-ttu-id="ee7a5-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="ee7a5-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee7a5-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ee7a5-187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee7a5-188">1.0</span><span class="sxs-lookup"><span data-stu-id="ee7a5-188">1.0</span></span>|
|[<span data-ttu-id="ee7a5-189">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ee7a5-189">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee7a5-190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee7a5-190">ReadItem</span></span>|
|[<span data-ttu-id="ee7a5-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ee7a5-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ee7a5-192">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ee7a5-192">Compose or Read</span></span>|
