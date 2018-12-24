---
title: Office.context.mailbox.userProfile- prévisualisations d’ensemble de conditions requises
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: ae9fca6ea2f7de99b275989bb2129948a60bc86f
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432997"
---
# <a name="diagnostics"></a><span data-ttu-id="46975-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="46975-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="46975-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="46975-103">Office.context.mailbox.diagnostics</span></span>

<span data-ttu-id="46975-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="46975-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="46975-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46975-105">Requirements</span></span>

|<span data-ttu-id="46975-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46975-106">Requirement</span></span>| <span data-ttu-id="46975-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="46975-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="46975-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46975-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46975-109">1.0</span><span class="sxs-lookup"><span data-stu-id="46975-109">1.0</span></span>|
|[<span data-ttu-id="46975-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46975-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46975-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46975-111">ReadItem</span></span>|
|[<span data-ttu-id="46975-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46975-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46975-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46975-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="46975-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="46975-114">Members and methods</span></span>

| <span data-ttu-id="46975-115">Membre</span><span class="sxs-lookup"><span data-stu-id="46975-115">Member</span></span> | <span data-ttu-id="46975-116">Type</span><span class="sxs-lookup"><span data-stu-id="46975-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="46975-117">hostName</span><span class="sxs-lookup"><span data-stu-id="46975-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="46975-118">Membre</span><span class="sxs-lookup"><span data-stu-id="46975-118">Member</span></span> |
| [<span data-ttu-id="46975-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="46975-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="46975-120">Membre</span><span class="sxs-lookup"><span data-stu-id="46975-120">Member</span></span> |
| [<span data-ttu-id="46975-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="46975-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="46975-122">Membre</span><span class="sxs-lookup"><span data-stu-id="46975-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="46975-123">Membres</span><span class="sxs-lookup"><span data-stu-id="46975-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="46975-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="46975-124">hostName :String</span></span>

<span data-ttu-id="46975-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="46975-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="46975-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="46975-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="46975-127">Type :</span><span class="sxs-lookup"><span data-stu-id="46975-127">Type:</span></span>

*   <span data-ttu-id="46975-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="46975-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="46975-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46975-129">Requirements</span></span>

|<span data-ttu-id="46975-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46975-130">Requirement</span></span>| <span data-ttu-id="46975-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="46975-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="46975-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46975-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46975-133">1.0</span><span class="sxs-lookup"><span data-stu-id="46975-133">1.0</span></span>|
|[<span data-ttu-id="46975-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46975-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46975-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46975-135">ReadItem</span></span>|
|[<span data-ttu-id="46975-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46975-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46975-137">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46975-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="46975-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="46975-138">hostVersion :String</span></span>

<span data-ttu-id="46975-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="46975-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="46975-p101">Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="46975-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="46975-143">Type :</span><span class="sxs-lookup"><span data-stu-id="46975-143">Type:</span></span>

*   <span data-ttu-id="46975-144">Chaîne</span><span class="sxs-lookup"><span data-stu-id="46975-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="46975-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46975-145">Requirements</span></span>

|<span data-ttu-id="46975-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46975-146">Requirement</span></span>| <span data-ttu-id="46975-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="46975-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="46975-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46975-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46975-149">1.0</span><span class="sxs-lookup"><span data-stu-id="46975-149">1.0</span></span>|
|[<span data-ttu-id="46975-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46975-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46975-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46975-151">ReadItem</span></span>|
|[<span data-ttu-id="46975-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46975-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46975-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46975-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="46975-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="46975-154">OWAView :String</span></span>

<span data-ttu-id="46975-155">Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="46975-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="46975-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="46975-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="46975-157">Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété génère la valeur `undefined`.</span><span class="sxs-lookup"><span data-stu-id="46975-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="46975-158">Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="46975-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="46975-p102">`OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="46975-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="46975-p103">`TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="46975-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="46975-p104">`ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode Plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="46975-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="46975-165">Type :</span><span class="sxs-lookup"><span data-stu-id="46975-165">Type:</span></span>

*   <span data-ttu-id="46975-166">Chaîne</span><span class="sxs-lookup"><span data-stu-id="46975-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="46975-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="46975-167">Requirements</span></span>

|<span data-ttu-id="46975-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="46975-168">Requirement</span></span>| <span data-ttu-id="46975-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="46975-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="46975-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="46975-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46975-171">1.0</span><span class="sxs-lookup"><span data-stu-id="46975-171">1.0</span></span>|
|[<span data-ttu-id="46975-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="46975-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="46975-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="46975-173">ReadItem</span></span>|
|[<span data-ttu-id="46975-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="46975-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46975-175">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="46975-175">Compose or read</span></span>|