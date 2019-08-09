---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,7
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 65fcaf2d7d04f56703ea6138d2d7820a34e5821c
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268319"
---
# <a name="diagnostics"></a><span data-ttu-id="bac40-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="bac40-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="bac40-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="bac40-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="bac40-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="bac40-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac40-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bac40-105">Requirements</span></span>

|<span data-ttu-id="bac40-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bac40-106">Requirement</span></span>| <span data-ttu-id="bac40-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="bac40-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac40-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bac40-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac40-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bac40-109">1.0</span></span>|
|[<span data-ttu-id="bac40-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bac40-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac40-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac40-111">ReadItem</span></span>|
|[<span data-ttu-id="bac40-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bac40-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bac40-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bac40-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bac40-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="bac40-114">Members and methods</span></span>

| <span data-ttu-id="bac40-115">Membre</span><span class="sxs-lookup"><span data-stu-id="bac40-115">Member</span></span> | <span data-ttu-id="bac40-116">Type</span><span class="sxs-lookup"><span data-stu-id="bac40-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bac40-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="bac40-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="bac40-118">Member</span><span class="sxs-lookup"><span data-stu-id="bac40-118">Member</span></span> |
| [<span data-ttu-id="bac40-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="bac40-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="bac40-120">Member</span><span class="sxs-lookup"><span data-stu-id="bac40-120">Member</span></span> |
| [<span data-ttu-id="bac40-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="bac40-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="bac40-122">Membre</span><span class="sxs-lookup"><span data-stu-id="bac40-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="bac40-123">Membres</span><span class="sxs-lookup"><span data-stu-id="bac40-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="bac40-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="bac40-124">hostName: String</span></span>

<span data-ttu-id="bac40-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="bac40-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="bac40-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="bac40-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="bac40-127">La `Outlook` valeur est renvoyée pour Outlook sur les clients de bureau (par exemple, Windows et Mac).</span><span class="sxs-lookup"><span data-stu-id="bac40-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="bac40-128">Type</span><span class="sxs-lookup"><span data-stu-id="bac40-128">Type</span></span>

*   <span data-ttu-id="bac40-129">String</span><span class="sxs-lookup"><span data-stu-id="bac40-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac40-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bac40-130">Requirements</span></span>

|<span data-ttu-id="bac40-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bac40-131">Requirement</span></span>| <span data-ttu-id="bac40-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="bac40-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac40-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bac40-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac40-134">1.0</span><span class="sxs-lookup"><span data-stu-id="bac40-134">1.0</span></span>|
|[<span data-ttu-id="bac40-135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bac40-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac40-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac40-136">ReadItem</span></span>|
|[<span data-ttu-id="bac40-137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bac40-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bac40-138">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bac40-138">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="bac40-139">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="bac40-139">hostVersion: String</span></span>

<span data-ttu-id="bac40-140">Obtient une valeur de type String qui représente la version de l’application hôte ou du serveur Exchange (par exemple, «15.0.468.0»).</span><span class="sxs-lookup"><span data-stu-id="bac40-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="bac40-141">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="bac40-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="bac40-142">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="bac40-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="bac40-143">Type</span><span class="sxs-lookup"><span data-stu-id="bac40-143">Type</span></span>

*   <span data-ttu-id="bac40-144">String</span><span class="sxs-lookup"><span data-stu-id="bac40-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac40-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bac40-145">Requirements</span></span>

|<span data-ttu-id="bac40-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bac40-146">Requirement</span></span>| <span data-ttu-id="bac40-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="bac40-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac40-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bac40-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac40-149">1.0</span><span class="sxs-lookup"><span data-stu-id="bac40-149">1.0</span></span>|
|[<span data-ttu-id="bac40-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bac40-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac40-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac40-151">ReadItem</span></span>|
|[<span data-ttu-id="bac40-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bac40-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bac40-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bac40-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="bac40-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="bac40-154">OWAView: String</span></span>

<span data-ttu-id="bac40-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="bac40-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="bac40-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="bac40-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="bac40-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bac40-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="bac40-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="bac40-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="bac40-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="bac40-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="bac40-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="bac40-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="bac40-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="bac40-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="bac40-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="bac40-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="bac40-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="bac40-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="bac40-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="bac40-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="bac40-165">Type</span><span class="sxs-lookup"><span data-stu-id="bac40-165">Type</span></span>

*   <span data-ttu-id="bac40-166">String</span><span class="sxs-lookup"><span data-stu-id="bac40-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac40-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bac40-167">Requirements</span></span>

|<span data-ttu-id="bac40-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bac40-168">Requirement</span></span>| <span data-ttu-id="bac40-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="bac40-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac40-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bac40-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac40-171">1.0</span><span class="sxs-lookup"><span data-stu-id="bac40-171">1.0</span></span>|
|[<span data-ttu-id="bac40-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bac40-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac40-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac40-173">ReadItem</span></span>|
|[<span data-ttu-id="bac40-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bac40-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bac40-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bac40-175">Compose or Read</span></span>|
