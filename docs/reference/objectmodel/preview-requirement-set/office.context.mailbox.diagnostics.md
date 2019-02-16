---
title: Office.context.mailbox.userProfile- prévisualisations d’ensemble de conditions requises
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: edeb0ce8a19dae4ec12705f81c216d4d7857ae90
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068167"
---
# <a name="diagnostics"></a><span data-ttu-id="bd671-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="bd671-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="bd671-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="bd671-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="bd671-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd671-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd671-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bd671-105">Requirements</span></span>

|<span data-ttu-id="bd671-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bd671-106">Requirement</span></span>| <span data-ttu-id="bd671-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="bd671-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd671-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bd671-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd671-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bd671-109">1.0</span></span>|
|[<span data-ttu-id="bd671-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bd671-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd671-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd671-111">ReadItem</span></span>|
|[<span data-ttu-id="bd671-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bd671-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bd671-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bd671-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bd671-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="bd671-114">Members and methods</span></span>

| <span data-ttu-id="bd671-115">Membre</span><span class="sxs-lookup"><span data-stu-id="bd671-115">Member</span></span> | <span data-ttu-id="bd671-116">Type</span><span class="sxs-lookup"><span data-stu-id="bd671-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bd671-117">hostName</span><span class="sxs-lookup"><span data-stu-id="bd671-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="bd671-118">Membre</span><span class="sxs-lookup"><span data-stu-id="bd671-118">Member</span></span> |
| [<span data-ttu-id="bd671-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="bd671-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="bd671-120">Membre</span><span class="sxs-lookup"><span data-stu-id="bd671-120">Member</span></span> |
| [<span data-ttu-id="bd671-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="bd671-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="bd671-122">Membre</span><span class="sxs-lookup"><span data-stu-id="bd671-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="bd671-123">Membres</span><span class="sxs-lookup"><span data-stu-id="bd671-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="bd671-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="bd671-124">hostName :String</span></span>

<span data-ttu-id="bd671-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="bd671-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="bd671-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="bd671-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="bd671-127">Type</span><span class="sxs-lookup"><span data-stu-id="bd671-127">Type</span></span>

*   <span data-ttu-id="bd671-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bd671-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd671-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bd671-129">Requirements</span></span>

|<span data-ttu-id="bd671-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bd671-130">Requirement</span></span>| <span data-ttu-id="bd671-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="bd671-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd671-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bd671-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd671-133">1.0</span><span class="sxs-lookup"><span data-stu-id="bd671-133">1.0</span></span>|
|[<span data-ttu-id="bd671-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bd671-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd671-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd671-135">ReadItem</span></span>|
|[<span data-ttu-id="bd671-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bd671-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bd671-137">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bd671-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="bd671-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="bd671-138">hostVersion :String</span></span>

<span data-ttu-id="bd671-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="bd671-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="bd671-p101">Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="bd671-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="bd671-143">Type</span><span class="sxs-lookup"><span data-stu-id="bd671-143">Type</span></span>

*   <span data-ttu-id="bd671-144">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bd671-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd671-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bd671-145">Requirements</span></span>

|<span data-ttu-id="bd671-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bd671-146">Requirement</span></span>| <span data-ttu-id="bd671-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="bd671-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd671-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bd671-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd671-149">1.0</span><span class="sxs-lookup"><span data-stu-id="bd671-149">1.0</span></span>|
|[<span data-ttu-id="bd671-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bd671-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd671-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd671-151">ReadItem</span></span>|
|[<span data-ttu-id="bd671-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bd671-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bd671-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bd671-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="bd671-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="bd671-154">OWAView :String</span></span>

<span data-ttu-id="bd671-155">Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="bd671-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="bd671-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="bd671-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="bd671-157">Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété génère la valeur `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bd671-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="bd671-158">Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="bd671-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="bd671-p102">`OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="bd671-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="bd671-p103">`TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="bd671-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="bd671-p104">`ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode Plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="bd671-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="bd671-165">Type</span><span class="sxs-lookup"><span data-stu-id="bd671-165">Type</span></span>

*   <span data-ttu-id="bd671-166">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bd671-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd671-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bd671-167">Requirements</span></span>

|<span data-ttu-id="bd671-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bd671-168">Requirement</span></span>| <span data-ttu-id="bd671-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="bd671-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd671-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bd671-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd671-171">1.0</span><span class="sxs-lookup"><span data-stu-id="bd671-171">1.0</span></span>|
|[<span data-ttu-id="bd671-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bd671-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd671-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd671-173">ReadItem</span></span>|
|[<span data-ttu-id="bd671-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bd671-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bd671-175">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bd671-175">Compose or Read</span></span>|
