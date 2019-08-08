---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,7
description: ''
ms.date: 08/05/2019
localization_priority: Normal
ms.openlocfilehash: e197374267d40056c6cb1dea8808e30f48eef65c
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231262"
---
# <a name="diagnostics"></a><span data-ttu-id="c987e-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="c987e-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="c987e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="c987e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="c987e-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c987e-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c987e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c987e-105">Requirements</span></span>

|<span data-ttu-id="c987e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c987e-106">Requirement</span></span>| <span data-ttu-id="c987e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c987e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c987e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c987e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c987e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c987e-109">1.0</span></span>|
|[<span data-ttu-id="c987e-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c987e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c987e-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c987e-111">ReadItem</span></span>|
|[<span data-ttu-id="c987e-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c987e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c987e-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c987e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c987e-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c987e-114">Members and methods</span></span>

| <span data-ttu-id="c987e-115">Membre</span><span class="sxs-lookup"><span data-stu-id="c987e-115">Member</span></span> | <span data-ttu-id="c987e-116">Type</span><span class="sxs-lookup"><span data-stu-id="c987e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c987e-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="c987e-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="c987e-118">Member</span><span class="sxs-lookup"><span data-stu-id="c987e-118">Member</span></span> |
| [<span data-ttu-id="c987e-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="c987e-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="c987e-120">Member</span><span class="sxs-lookup"><span data-stu-id="c987e-120">Member</span></span> |
| [<span data-ttu-id="c987e-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="c987e-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="c987e-122">Membre</span><span class="sxs-lookup"><span data-stu-id="c987e-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c987e-123">Membres</span><span class="sxs-lookup"><span data-stu-id="c987e-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="c987e-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="c987e-124">hostName: String</span></span>

<span data-ttu-id="c987e-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="c987e-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="c987e-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="c987e-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

##### <a name="type"></a><span data-ttu-id="c987e-127">Type</span><span class="sxs-lookup"><span data-stu-id="c987e-127">Type</span></span>

*   <span data-ttu-id="c987e-128">String</span><span class="sxs-lookup"><span data-stu-id="c987e-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c987e-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c987e-129">Requirements</span></span>

|<span data-ttu-id="c987e-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c987e-130">Requirement</span></span>| <span data-ttu-id="c987e-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="c987e-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="c987e-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c987e-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c987e-133">1.0</span><span class="sxs-lookup"><span data-stu-id="c987e-133">1.0</span></span>|
|[<span data-ttu-id="c987e-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c987e-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c987e-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c987e-135">ReadItem</span></span>|
|[<span data-ttu-id="c987e-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c987e-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c987e-137">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c987e-137">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="c987e-138">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="c987e-138">hostVersion: String</span></span>

<span data-ttu-id="c987e-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="c987e-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="c987e-140">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="c987e-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="c987e-141">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="c987e-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="c987e-142">La chaîne «15.0.468.0» est un exemple.</span><span class="sxs-lookup"><span data-stu-id="c987e-142">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="c987e-143">Type</span><span class="sxs-lookup"><span data-stu-id="c987e-143">Type</span></span>

*   <span data-ttu-id="c987e-144">String</span><span class="sxs-lookup"><span data-stu-id="c987e-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c987e-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c987e-145">Requirements</span></span>

|<span data-ttu-id="c987e-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c987e-146">Requirement</span></span>| <span data-ttu-id="c987e-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="c987e-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="c987e-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c987e-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c987e-149">1.0</span><span class="sxs-lookup"><span data-stu-id="c987e-149">1.0</span></span>|
|[<span data-ttu-id="c987e-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c987e-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c987e-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c987e-151">ReadItem</span></span>|
|[<span data-ttu-id="c987e-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c987e-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c987e-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c987e-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="c987e-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="c987e-154">OWAView: String</span></span>

<span data-ttu-id="c987e-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="c987e-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="c987e-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="c987e-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="c987e-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c987e-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="c987e-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="c987e-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="c987e-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="c987e-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="c987e-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="c987e-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="c987e-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="c987e-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="c987e-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="c987e-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="c987e-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="c987e-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="c987e-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="c987e-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="c987e-165">Type</span><span class="sxs-lookup"><span data-stu-id="c987e-165">Type</span></span>

*   <span data-ttu-id="c987e-166">String</span><span class="sxs-lookup"><span data-stu-id="c987e-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c987e-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c987e-167">Requirements</span></span>

|<span data-ttu-id="c987e-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c987e-168">Requirement</span></span>| <span data-ttu-id="c987e-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="c987e-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="c987e-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c987e-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c987e-171">1.0</span><span class="sxs-lookup"><span data-stu-id="c987e-171">1.0</span></span>|
|[<span data-ttu-id="c987e-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c987e-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c987e-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c987e-173">ReadItem</span></span>|
|[<span data-ttu-id="c987e-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c987e-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c987e-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c987e-175">Compose or Read</span></span>|
