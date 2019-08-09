---
title: Office.context.mailbox.diagnostics – ensemble de conditions requises 1.1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: a2d54aee0545fb43aea798d799f190519226d185
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268508"
---
# <a name="diagnostics"></a><span data-ttu-id="14c4d-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="14c4d-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="14c4d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="14c4d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="14c4d-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="14c4d-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c4d-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14c4d-105">Requirements</span></span>

|<span data-ttu-id="14c4d-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14c4d-106">Requirement</span></span>| <span data-ttu-id="14c4d-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="14c4d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c4d-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14c4d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c4d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="14c4d-109">1.0</span></span>|
|[<span data-ttu-id="14c4d-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14c4d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c4d-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c4d-111">ReadItem</span></span>|
|[<span data-ttu-id="14c4d-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14c4d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c4d-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14c4d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="14c4d-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="14c4d-114">Members and methods</span></span>

| <span data-ttu-id="14c4d-115">Membre</span><span class="sxs-lookup"><span data-stu-id="14c4d-115">Member</span></span> | <span data-ttu-id="14c4d-116">Type</span><span class="sxs-lookup"><span data-stu-id="14c4d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="14c4d-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="14c4d-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="14c4d-118">Member</span><span class="sxs-lookup"><span data-stu-id="14c4d-118">Member</span></span> |
| [<span data-ttu-id="14c4d-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="14c4d-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="14c4d-120">Member</span><span class="sxs-lookup"><span data-stu-id="14c4d-120">Member</span></span> |
| [<span data-ttu-id="14c4d-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="14c4d-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="14c4d-122">Membre</span><span class="sxs-lookup"><span data-stu-id="14c4d-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="14c4d-123">Membres</span><span class="sxs-lookup"><span data-stu-id="14c4d-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="14c4d-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="14c4d-124">hostName: String</span></span>

<span data-ttu-id="14c4d-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="14c4d-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="14c4d-126">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="14c4d-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

> [!NOTE]
> <span data-ttu-id="14c4d-127">La `Outlook` valeur est renvoyée pour Outlook sur les clients de bureau (par exemple, Windows et Mac).</span><span class="sxs-lookup"><span data-stu-id="14c4d-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="14c4d-128">Type</span><span class="sxs-lookup"><span data-stu-id="14c4d-128">Type</span></span>

*   <span data-ttu-id="14c4d-129">String</span><span class="sxs-lookup"><span data-stu-id="14c4d-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c4d-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14c4d-130">Requirements</span></span>

|<span data-ttu-id="14c4d-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14c4d-131">Requirement</span></span>| <span data-ttu-id="14c4d-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="14c4d-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c4d-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14c4d-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c4d-134">1.0</span><span class="sxs-lookup"><span data-stu-id="14c4d-134">1.0</span></span>|
|[<span data-ttu-id="14c4d-135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14c4d-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c4d-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c4d-136">ReadItem</span></span>|
|[<span data-ttu-id="14c4d-137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14c4d-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c4d-138">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14c4d-138">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="14c4d-139">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="14c4d-139">hostVersion: String</span></span>

<span data-ttu-id="14c4d-140">Obtient une valeur de type String qui représente la version de l’application hôte ou du serveur Exchange (par exemple, «15.0.468.0»).</span><span class="sxs-lookup"><span data-stu-id="14c4d-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="14c4d-141">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="14c4d-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="14c4d-142">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="14c4d-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="14c4d-143">Type</span><span class="sxs-lookup"><span data-stu-id="14c4d-143">Type</span></span>

*   <span data-ttu-id="14c4d-144">String</span><span class="sxs-lookup"><span data-stu-id="14c4d-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c4d-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14c4d-145">Requirements</span></span>

|<span data-ttu-id="14c4d-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14c4d-146">Requirement</span></span>| <span data-ttu-id="14c4d-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="14c4d-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c4d-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14c4d-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c4d-149">1.0</span><span class="sxs-lookup"><span data-stu-id="14c4d-149">1.0</span></span>|
|[<span data-ttu-id="14c4d-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14c4d-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c4d-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c4d-151">ReadItem</span></span>|
|[<span data-ttu-id="14c4d-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14c4d-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c4d-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14c4d-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="14c4d-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="14c4d-154">OWAView: String</span></span>

<span data-ttu-id="14c4d-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="14c4d-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="14c4d-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="14c4d-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="14c4d-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14c4d-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="14c4d-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="14c4d-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="14c4d-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="14c4d-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="14c4d-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="14c4d-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="14c4d-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="14c4d-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="14c4d-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="14c4d-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="14c4d-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="14c4d-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="14c4d-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="14c4d-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="14c4d-165">Type</span><span class="sxs-lookup"><span data-stu-id="14c4d-165">Type</span></span>

*   <span data-ttu-id="14c4d-166">String</span><span class="sxs-lookup"><span data-stu-id="14c4d-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c4d-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="14c4d-167">Requirements</span></span>

|<span data-ttu-id="14c4d-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="14c4d-168">Requirement</span></span>| <span data-ttu-id="14c4d-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="14c4d-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c4d-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="14c4d-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c4d-171">1.0</span><span class="sxs-lookup"><span data-stu-id="14c4d-171">1.0</span></span>|
|[<span data-ttu-id="14c4d-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="14c4d-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c4d-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c4d-173">ReadItem</span></span>|
|[<span data-ttu-id="14c4d-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="14c4d-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c4d-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="14c4d-175">Compose or Read</span></span>|
