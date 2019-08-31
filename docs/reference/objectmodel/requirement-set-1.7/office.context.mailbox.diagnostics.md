---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,7
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 3f17f21bd4fcf48496ba7f387a57752256d4a038
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695965"
---
# <a name="diagnostics"></a><span data-ttu-id="bc5a3-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="bc5a3-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="bc5a3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="bc5a3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="bc5a3-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc5a3-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc5a3-105">Requirements</span></span>

|<span data-ttu-id="bc5a3-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc5a3-106">Requirement</span></span>| <span data-ttu-id="bc5a3-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc5a3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc5a3-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc5a3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc5a3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bc5a3-109">1.0</span></span>|
|[<span data-ttu-id="bc5a3-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc5a3-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc5a3-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc5a3-111">ReadItem</span></span>|
|[<span data-ttu-id="bc5a3-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc5a3-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc5a3-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc5a3-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bc5a3-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="bc5a3-114">Members and methods</span></span>

| <span data-ttu-id="bc5a3-115">Membre</span><span class="sxs-lookup"><span data-stu-id="bc5a3-115">Member</span></span> | <span data-ttu-id="bc5a3-116">Type</span><span class="sxs-lookup"><span data-stu-id="bc5a3-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bc5a3-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="bc5a3-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="bc5a3-118">Member</span><span class="sxs-lookup"><span data-stu-id="bc5a3-118">Member</span></span> |
| [<span data-ttu-id="bc5a3-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="bc5a3-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="bc5a3-120">Member</span><span class="sxs-lookup"><span data-stu-id="bc5a3-120">Member</span></span> |
| [<span data-ttu-id="bc5a3-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="bc5a3-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="bc5a3-122">Membre</span><span class="sxs-lookup"><span data-stu-id="bc5a3-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="bc5a3-123">Membres</span><span class="sxs-lookup"><span data-stu-id="bc5a3-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="bc5a3-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="bc5a3-124">hostName: String</span></span>

<span data-ttu-id="bc5a3-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="bc5a3-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="bc5a3-127">La `Outlook` valeur est renvoyée pour Outlook sur les clients de bureau (par exemple, Windows et Mac).</span><span class="sxs-lookup"><span data-stu-id="bc5a3-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="bc5a3-128">Type</span><span class="sxs-lookup"><span data-stu-id="bc5a3-128">Type</span></span>

*   <span data-ttu-id="bc5a3-129">String</span><span class="sxs-lookup"><span data-stu-id="bc5a3-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc5a3-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc5a3-130">Requirements</span></span>

|<span data-ttu-id="bc5a3-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc5a3-131">Requirement</span></span>| <span data-ttu-id="bc5a3-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc5a3-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc5a3-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc5a3-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc5a3-134">1.0</span><span class="sxs-lookup"><span data-stu-id="bc5a3-134">1.0</span></span>|
|[<span data-ttu-id="bc5a3-135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc5a3-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc5a3-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc5a3-136">ReadItem</span></span>|
|[<span data-ttu-id="bc5a3-137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc5a3-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc5a3-138">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc5a3-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="bc5a3-139">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="bc5a3-139">hostVersion: String</span></span>

<span data-ttu-id="bc5a3-140">Obtient une valeur de type String qui représente la version de l’application hôte ou du serveur Exchange (par exemple, «15.0.468.0»).</span><span class="sxs-lookup"><span data-stu-id="bc5a3-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="bc5a3-141">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="bc5a3-142">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="bc5a3-143">Type</span><span class="sxs-lookup"><span data-stu-id="bc5a3-143">Type</span></span>

*   <span data-ttu-id="bc5a3-144">String</span><span class="sxs-lookup"><span data-stu-id="bc5a3-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc5a3-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc5a3-145">Requirements</span></span>

|<span data-ttu-id="bc5a3-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc5a3-146">Requirement</span></span>| <span data-ttu-id="bc5a3-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc5a3-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc5a3-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc5a3-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc5a3-149">1.0</span><span class="sxs-lookup"><span data-stu-id="bc5a3-149">1.0</span></span>|
|[<span data-ttu-id="bc5a3-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc5a3-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc5a3-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc5a3-151">ReadItem</span></span>|
|[<span data-ttu-id="bc5a3-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc5a3-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc5a3-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc5a3-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="bc5a3-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="bc5a3-154">OWAView: String</span></span>

<span data-ttu-id="bc5a3-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="bc5a3-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="bc5a3-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="bc5a3-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="bc5a3-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="bc5a3-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="bc5a3-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="bc5a3-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="bc5a3-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="bc5a3-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="bc5a3-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="bc5a3-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="bc5a3-165">Type</span><span class="sxs-lookup"><span data-stu-id="bc5a3-165">Type</span></span>

*   <span data-ttu-id="bc5a3-166">String</span><span class="sxs-lookup"><span data-stu-id="bc5a3-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc5a3-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc5a3-167">Requirements</span></span>

|<span data-ttu-id="bc5a3-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc5a3-168">Requirement</span></span>| <span data-ttu-id="bc5a3-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc5a3-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc5a3-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc5a3-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc5a3-171">1.0</span><span class="sxs-lookup"><span data-stu-id="bc5a3-171">1.0</span></span>|
|[<span data-ttu-id="bc5a3-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc5a3-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc5a3-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc5a3-173">ReadItem</span></span>|
|[<span data-ttu-id="bc5a3-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc5a3-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc5a3-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc5a3-175">Compose or Read</span></span>|
