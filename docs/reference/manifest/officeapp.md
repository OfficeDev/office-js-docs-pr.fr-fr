---
title: Élément OfficeApp dans le fichier manifeste
description: L’élément OfficeApp est l’élément racine d’un manifeste de complément Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 770c764db6d8d7d1d2e870e48437de7c8f887101
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641458"
---
# <a name="officeapp-element"></a><span data-ttu-id="6cccd-103">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="6cccd-103">OfficeApp element</span></span>

<span data-ttu-id="6cccd-104">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="6cccd-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="6cccd-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="6cccd-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6cccd-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6cccd-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="6cccd-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="6cccd-107">Contained in</span></span>

 <span data-ttu-id="6cccd-108">_none_</span><span class="sxs-lookup"><span data-stu-id="6cccd-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="6cccd-109">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="6cccd-109">Must contain</span></span>

|<span data-ttu-id="6cccd-110">Élément</span><span class="sxs-lookup"><span data-stu-id="6cccd-110">Element</span></span>|<span data-ttu-id="6cccd-111">Contenu</span><span class="sxs-lookup"><span data-stu-id="6cccd-111">Content</span></span>|<span data-ttu-id="6cccd-112">Courrier</span><span class="sxs-lookup"><span data-stu-id="6cccd-112">Mail</span></span>|<span data-ttu-id="6cccd-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="6cccd-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="6cccd-114">Id</span><span class="sxs-lookup"><span data-stu-id="6cccd-114">Id</span></span>](id.md)|<span data-ttu-id="6cccd-115">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-115">x</span></span>|<span data-ttu-id="6cccd-116">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-116">x</span></span>|<span data-ttu-id="6cccd-117">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-117">x</span></span>|
|[<span data-ttu-id="6cccd-118">Version</span><span class="sxs-lookup"><span data-stu-id="6cccd-118">Version</span></span>](version.md)|<span data-ttu-id="6cccd-119">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-119">x</span></span>|<span data-ttu-id="6cccd-120">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-120">x</span></span>|<span data-ttu-id="6cccd-121">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-121">x</span></span>|
|[<span data-ttu-id="6cccd-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="6cccd-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="6cccd-123">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-123">x</span></span>|<span data-ttu-id="6cccd-124">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-124">x</span></span>|<span data-ttu-id="6cccd-125">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-125">x</span></span>|
|[<span data-ttu-id="6cccd-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="6cccd-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="6cccd-127">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-127">x</span></span>|<span data-ttu-id="6cccd-128">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-128">x</span></span>|<span data-ttu-id="6cccd-129">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-129">x</span></span>|
|[<span data-ttu-id="6cccd-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="6cccd-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="6cccd-131">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-131">x</span></span>||<span data-ttu-id="6cccd-132">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-132">x</span></span>|
|[<span data-ttu-id="6cccd-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="6cccd-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="6cccd-134">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-134">x</span></span>|<span data-ttu-id="6cccd-135">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-135">x</span></span>|<span data-ttu-id="6cccd-136">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-136">x</span></span>|
|[<span data-ttu-id="6cccd-137">Description</span><span class="sxs-lookup"><span data-stu-id="6cccd-137">Description</span></span>](description.md)|<span data-ttu-id="6cccd-138">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-138">x</span></span>|<span data-ttu-id="6cccd-139">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-139">x</span></span>|<span data-ttu-id="6cccd-140">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-140">x</span></span>|
|[<span data-ttu-id="6cccd-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="6cccd-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="6cccd-142">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-142">x</span></span>||
|[<span data-ttu-id="6cccd-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="6cccd-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="6cccd-144">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-144">x</span></span>||<span data-ttu-id="6cccd-145">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-145">x</span></span>|
|[<span data-ttu-id="6cccd-146">Règle</span><span class="sxs-lookup"><span data-stu-id="6cccd-146">Rule</span></span>](rule.md)||<span data-ttu-id="6cccd-147">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="6cccd-148">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="6cccd-148">Can contain</span></span>

|<span data-ttu-id="6cccd-149">Élément</span><span class="sxs-lookup"><span data-stu-id="6cccd-149">Element</span></span>|<span data-ttu-id="6cccd-150">Contenu</span><span class="sxs-lookup"><span data-stu-id="6cccd-150">Content</span></span>|<span data-ttu-id="6cccd-151">Courrier</span><span class="sxs-lookup"><span data-stu-id="6cccd-151">Mail</span></span>|<span data-ttu-id="6cccd-152">TaskPane</span><span class="sxs-lookup"><span data-stu-id="6cccd-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="6cccd-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="6cccd-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="6cccd-154">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-154">x</span></span>|<span data-ttu-id="6cccd-155">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-155">x</span></span>|<span data-ttu-id="6cccd-156">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-156">x</span></span>|
|[<span data-ttu-id="6cccd-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="6cccd-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="6cccd-158">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-158">x</span></span>|<span data-ttu-id="6cccd-159">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-159">x</span></span>|<span data-ttu-id="6cccd-160">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-160">x</span></span>|
|[<span data-ttu-id="6cccd-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="6cccd-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="6cccd-162">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-162">x</span></span>|<span data-ttu-id="6cccd-163">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-163">x</span></span>|<span data-ttu-id="6cccd-164">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-164">x</span></span>|
|[<span data-ttu-id="6cccd-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="6cccd-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="6cccd-166">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-166">x</span></span>|<span data-ttu-id="6cccd-167">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-167">x</span></span>|<span data-ttu-id="6cccd-168">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-168">x</span></span>|
|[<span data-ttu-id="6cccd-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="6cccd-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="6cccd-170">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-170">x</span></span>|<span data-ttu-id="6cccd-171">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-171">x</span></span>|<span data-ttu-id="6cccd-172">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-172">x</span></span>|
|[<span data-ttu-id="6cccd-173">Hôtes</span><span class="sxs-lookup"><span data-stu-id="6cccd-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="6cccd-174">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-174">x</span></span>|<span data-ttu-id="6cccd-175">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-175">x</span></span>|<span data-ttu-id="6cccd-176">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-176">x</span></span>|
|[<span data-ttu-id="6cccd-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cccd-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="6cccd-178">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-178">x</span></span>|<span data-ttu-id="6cccd-179">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-179">x</span></span>|<span data-ttu-id="6cccd-180">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-180">x</span></span>|
|[<span data-ttu-id="6cccd-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="6cccd-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="6cccd-182">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-182">x</span></span>|||
|[<span data-ttu-id="6cccd-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="6cccd-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="6cccd-184">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-184">x</span></span>||
|[<span data-ttu-id="6cccd-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="6cccd-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="6cccd-186">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-186">x</span></span>||
|[<span data-ttu-id="6cccd-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="6cccd-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="6cccd-188">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-188">x</span></span>|
|[<span data-ttu-id="6cccd-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="6cccd-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="6cccd-190">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-190">x</span></span>|<span data-ttu-id="6cccd-191">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-191">x</span></span>|<span data-ttu-id="6cccd-192">x</span><span class="sxs-lookup"><span data-stu-id="6cccd-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="6cccd-193">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cccd-193">Attributes</span></span>

|<span data-ttu-id="6cccd-194">Attribut</span><span class="sxs-lookup"><span data-stu-id="6cccd-194">Attribute</span></span>|<span data-ttu-id="6cccd-195">Description</span><span class="sxs-lookup"><span data-stu-id="6cccd-195">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="6cccd-196">xmlns</span><span class="sxs-lookup"><span data-stu-id="6cccd-196">xmlns</span></span>|<span data-ttu-id="6cccd-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="6cccd-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="6cccd-199">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="6cccd-199">xmlns:xsi</span></span>|<span data-ttu-id="6cccd-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="6cccd-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="6cccd-202">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6cccd-202">xsi:type</span></span>|<span data-ttu-id="6cccd-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="6cccd-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
