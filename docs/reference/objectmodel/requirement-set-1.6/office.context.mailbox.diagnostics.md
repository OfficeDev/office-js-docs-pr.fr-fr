---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,6
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 04fd4af8e35b2a538e93a64254250d40c3334dc6
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127372"
---
# <a name="diagnostics"></a><span data-ttu-id="444a9-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="444a9-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="444a9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="444a9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="444a9-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="444a9-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="444a9-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="444a9-105">Requirements</span></span>

|<span data-ttu-id="444a9-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="444a9-106">Requirement</span></span>| <span data-ttu-id="444a9-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="444a9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="444a9-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="444a9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="444a9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="444a9-109">1.0</span></span>|
|[<span data-ttu-id="444a9-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="444a9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="444a9-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="444a9-111">ReadItem</span></span>|
|[<span data-ttu-id="444a9-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="444a9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="444a9-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="444a9-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="444a9-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="444a9-114">Members and methods</span></span>

| <span data-ttu-id="444a9-115">Membre</span><span class="sxs-lookup"><span data-stu-id="444a9-115">Member</span></span> | <span data-ttu-id="444a9-116">Type</span><span class="sxs-lookup"><span data-stu-id="444a9-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="444a9-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="444a9-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="444a9-118">Member</span><span class="sxs-lookup"><span data-stu-id="444a9-118">Member</span></span> |
| [<span data-ttu-id="444a9-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="444a9-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="444a9-120">Member</span><span class="sxs-lookup"><span data-stu-id="444a9-120">Member</span></span> |
| [<span data-ttu-id="444a9-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="444a9-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="444a9-122">Membre</span><span class="sxs-lookup"><span data-stu-id="444a9-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="444a9-123">Membres</span><span class="sxs-lookup"><span data-stu-id="444a9-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="444a9-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="444a9-124">hostName: String</span></span>

<span data-ttu-id="444a9-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="444a9-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="444a9-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="444a9-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="444a9-127">Type</span><span class="sxs-lookup"><span data-stu-id="444a9-127">Type</span></span>

*   <span data-ttu-id="444a9-128">String</span><span class="sxs-lookup"><span data-stu-id="444a9-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="444a9-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="444a9-129">Requirements</span></span>

|<span data-ttu-id="444a9-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="444a9-130">Requirement</span></span>| <span data-ttu-id="444a9-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="444a9-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="444a9-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="444a9-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="444a9-133">1.0</span><span class="sxs-lookup"><span data-stu-id="444a9-133">1.0</span></span>|
|[<span data-ttu-id="444a9-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="444a9-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="444a9-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="444a9-135">ReadItem</span></span>|
|[<span data-ttu-id="444a9-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="444a9-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="444a9-137">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="444a9-137">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="444a9-138">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="444a9-138">hostVersion: String</span></span>

<span data-ttu-id="444a9-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="444a9-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="444a9-140">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="444a9-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="444a9-141">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="444a9-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="444a9-142">Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="444a9-142">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="444a9-143">Type</span><span class="sxs-lookup"><span data-stu-id="444a9-143">Type</span></span>

*   <span data-ttu-id="444a9-144">String</span><span class="sxs-lookup"><span data-stu-id="444a9-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="444a9-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="444a9-145">Requirements</span></span>

|<span data-ttu-id="444a9-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="444a9-146">Requirement</span></span>| <span data-ttu-id="444a9-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="444a9-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="444a9-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="444a9-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="444a9-149">1.0</span><span class="sxs-lookup"><span data-stu-id="444a9-149">1.0</span></span>|
|[<span data-ttu-id="444a9-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="444a9-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="444a9-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="444a9-151">ReadItem</span></span>|
|[<span data-ttu-id="444a9-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="444a9-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="444a9-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="444a9-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="444a9-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="444a9-154">OWAView: String</span></span>

<span data-ttu-id="444a9-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="444a9-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="444a9-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="444a9-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="444a9-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="444a9-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="444a9-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="444a9-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="444a9-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="444a9-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="444a9-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="444a9-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="444a9-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="444a9-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="444a9-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="444a9-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="444a9-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="444a9-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="444a9-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="444a9-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="444a9-165">Type</span><span class="sxs-lookup"><span data-stu-id="444a9-165">Type</span></span>

*   <span data-ttu-id="444a9-166">String</span><span class="sxs-lookup"><span data-stu-id="444a9-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="444a9-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="444a9-167">Requirements</span></span>

|<span data-ttu-id="444a9-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="444a9-168">Requirement</span></span>| <span data-ttu-id="444a9-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="444a9-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="444a9-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="444a9-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="444a9-171">1.0</span><span class="sxs-lookup"><span data-stu-id="444a9-171">1.0</span></span>|
|[<span data-ttu-id="444a9-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="444a9-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="444a9-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="444a9-173">ReadItem</span></span>|
|[<span data-ttu-id="444a9-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="444a9-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="444a9-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="444a9-175">Compose or Read</span></span>|
