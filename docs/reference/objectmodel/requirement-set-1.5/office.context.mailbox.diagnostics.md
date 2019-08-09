---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,5
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 1a31a859eb79625943c3a2191f77c91535418b5e
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268627"
---
# <a name="diagnostics"></a><span data-ttu-id="3c584-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="3c584-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="3c584-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="3c584-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="3c584-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="3c584-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c584-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3c584-105">Requirements</span></span>

|<span data-ttu-id="3c584-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3c584-106">Requirement</span></span>| <span data-ttu-id="3c584-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="3c584-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c584-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3c584-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c584-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3c584-109">1.0</span></span>|
|[<span data-ttu-id="3c584-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3c584-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c584-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c584-111">ReadItem</span></span>|
|[<span data-ttu-id="3c584-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3c584-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3c584-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3c584-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3c584-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="3c584-114">Members and methods</span></span>

| <span data-ttu-id="3c584-115">Membre</span><span class="sxs-lookup"><span data-stu-id="3c584-115">Member</span></span> | <span data-ttu-id="3c584-116">Type</span><span class="sxs-lookup"><span data-stu-id="3c584-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3c584-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="3c584-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="3c584-118">Member</span><span class="sxs-lookup"><span data-stu-id="3c584-118">Member</span></span> |
| [<span data-ttu-id="3c584-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="3c584-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="3c584-120">Member</span><span class="sxs-lookup"><span data-stu-id="3c584-120">Member</span></span> |
| [<span data-ttu-id="3c584-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="3c584-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="3c584-122">Membre</span><span class="sxs-lookup"><span data-stu-id="3c584-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="3c584-123">Membres</span><span class="sxs-lookup"><span data-stu-id="3c584-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="3c584-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="3c584-124">hostName: String</span></span>

<span data-ttu-id="3c584-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="3c584-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="3c584-126">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="3c584-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

> [!NOTE]
> <span data-ttu-id="3c584-127">La `Outlook` valeur est renvoyée pour Outlook sur les clients de bureau (par exemple, Windows et Mac).</span><span class="sxs-lookup"><span data-stu-id="3c584-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="3c584-128">Type</span><span class="sxs-lookup"><span data-stu-id="3c584-128">Type</span></span>

*   <span data-ttu-id="3c584-129">String</span><span class="sxs-lookup"><span data-stu-id="3c584-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c584-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3c584-130">Requirements</span></span>

|<span data-ttu-id="3c584-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3c584-131">Requirement</span></span>| <span data-ttu-id="3c584-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="3c584-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c584-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3c584-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c584-134">1.0</span><span class="sxs-lookup"><span data-stu-id="3c584-134">1.0</span></span>|
|[<span data-ttu-id="3c584-135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3c584-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c584-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c584-136">ReadItem</span></span>|
|[<span data-ttu-id="3c584-137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3c584-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3c584-138">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3c584-138">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="3c584-139">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="3c584-139">hostVersion: String</span></span>

<span data-ttu-id="3c584-140">Obtient une valeur de type String qui représente la version de l’application hôte ou du serveur Exchange (par exemple, «15.0.468.0»).</span><span class="sxs-lookup"><span data-stu-id="3c584-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="3c584-141">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="3c584-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="3c584-142">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="3c584-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="3c584-143">Type</span><span class="sxs-lookup"><span data-stu-id="3c584-143">Type</span></span>

*   <span data-ttu-id="3c584-144">String</span><span class="sxs-lookup"><span data-stu-id="3c584-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c584-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3c584-145">Requirements</span></span>

|<span data-ttu-id="3c584-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3c584-146">Requirement</span></span>| <span data-ttu-id="3c584-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="3c584-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c584-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3c584-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c584-149">1.0</span><span class="sxs-lookup"><span data-stu-id="3c584-149">1.0</span></span>|
|[<span data-ttu-id="3c584-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3c584-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c584-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c584-151">ReadItem</span></span>|
|[<span data-ttu-id="3c584-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3c584-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3c584-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3c584-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="3c584-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="3c584-154">OWAView: String</span></span>

<span data-ttu-id="3c584-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="3c584-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="3c584-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="3c584-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="3c584-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="3c584-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="3c584-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="3c584-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="3c584-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="3c584-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="3c584-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="3c584-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="3c584-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="3c584-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="3c584-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="3c584-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="3c584-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="3c584-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="3c584-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="3c584-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="3c584-165">Type</span><span class="sxs-lookup"><span data-stu-id="3c584-165">Type</span></span>

*   <span data-ttu-id="3c584-166">String</span><span class="sxs-lookup"><span data-stu-id="3c584-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c584-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3c584-167">Requirements</span></span>

|<span data-ttu-id="3c584-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3c584-168">Requirement</span></span>| <span data-ttu-id="3c584-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="3c584-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c584-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3c584-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c584-171">1.0</span><span class="sxs-lookup"><span data-stu-id="3c584-171">1.0</span></span>|
|[<span data-ttu-id="3c584-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="3c584-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c584-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c584-173">ReadItem</span></span>|
|[<span data-ttu-id="3c584-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3c584-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3c584-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3c584-175">Compose or Read</span></span>|
