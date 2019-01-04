---
title: Élément OfficeApp dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 42b6fe2e1c33322b90016d5e7ceec7b1bfe5b72d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433165"
---
# <a name="officeapp-element"></a><span data-ttu-id="337d8-102">OfficeApp, élément</span><span class="sxs-lookup"><span data-stu-id="337d8-102">OfficeApp element</span></span>

<span data-ttu-id="337d8-103">Élément racine dans le manifeste d’un complément Office.</span><span class="sxs-lookup"><span data-stu-id="337d8-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="337d8-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="337d8-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="337d8-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="337d8-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="337d8-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="337d8-106">Contained in</span></span>

 <span data-ttu-id="337d8-107">_none_</span><span class="sxs-lookup"><span data-stu-id="337d8-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="337d8-108">Doit contenir</span><span class="sxs-lookup"><span data-stu-id="337d8-108">Must contain</span></span>

|<span data-ttu-id="337d8-109">**Élément**</span><span class="sxs-lookup"><span data-stu-id="337d8-109">**Element**</span></span>|<span data-ttu-id="337d8-110">**Contenu**</span><span class="sxs-lookup"><span data-stu-id="337d8-110">**Content**</span></span>|<span data-ttu-id="337d8-111">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="337d8-111">**Mail**</span></span>|<span data-ttu-id="337d8-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="337d8-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="337d8-113">Id</span><span class="sxs-lookup"><span data-stu-id="337d8-113">Id</span></span>](id.md)|<span data-ttu-id="337d8-114">x</span><span class="sxs-lookup"><span data-stu-id="337d8-114">x</span></span>|<span data-ttu-id="337d8-115">x</span><span class="sxs-lookup"><span data-stu-id="337d8-115">x</span></span>|<span data-ttu-id="337d8-116">x</span><span class="sxs-lookup"><span data-stu-id="337d8-116">x</span></span>|
|[<span data-ttu-id="337d8-117">Version</span><span class="sxs-lookup"><span data-stu-id="337d8-117">Version</span></span>](version.md)|<span data-ttu-id="337d8-118">x</span><span class="sxs-lookup"><span data-stu-id="337d8-118">x</span></span>|<span data-ttu-id="337d8-119">x</span><span class="sxs-lookup"><span data-stu-id="337d8-119">x</span></span>|<span data-ttu-id="337d8-120">x</span><span class="sxs-lookup"><span data-stu-id="337d8-120">x</span></span>|
|[<span data-ttu-id="337d8-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="337d8-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="337d8-122">x</span><span class="sxs-lookup"><span data-stu-id="337d8-122">x</span></span>|<span data-ttu-id="337d8-123">x</span><span class="sxs-lookup"><span data-stu-id="337d8-123">x</span></span>|<span data-ttu-id="337d8-124">x</span><span class="sxs-lookup"><span data-stu-id="337d8-124">x</span></span>|
|[<span data-ttu-id="337d8-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="337d8-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="337d8-126">x</span><span class="sxs-lookup"><span data-stu-id="337d8-126">x</span></span>|<span data-ttu-id="337d8-127">x</span><span class="sxs-lookup"><span data-stu-id="337d8-127">x</span></span>|<span data-ttu-id="337d8-128">x</span><span class="sxs-lookup"><span data-stu-id="337d8-128">x</span></span>|
|[<span data-ttu-id="337d8-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="337d8-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="337d8-130">x</span><span class="sxs-lookup"><span data-stu-id="337d8-130">x</span></span>||<span data-ttu-id="337d8-131">x</span><span class="sxs-lookup"><span data-stu-id="337d8-131">x</span></span>|
|[<span data-ttu-id="337d8-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="337d8-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="337d8-133">x</span><span class="sxs-lookup"><span data-stu-id="337d8-133">x</span></span>|<span data-ttu-id="337d8-134">x</span><span class="sxs-lookup"><span data-stu-id="337d8-134">x</span></span>|<span data-ttu-id="337d8-135">x</span><span class="sxs-lookup"><span data-stu-id="337d8-135">x</span></span>|
|[<span data-ttu-id="337d8-136">Description</span><span class="sxs-lookup"><span data-stu-id="337d8-136">Description</span></span>](description.md)|<span data-ttu-id="337d8-137">x</span><span class="sxs-lookup"><span data-stu-id="337d8-137">x</span></span>|<span data-ttu-id="337d8-138">x</span><span class="sxs-lookup"><span data-stu-id="337d8-138">x</span></span>|<span data-ttu-id="337d8-139">x</span><span class="sxs-lookup"><span data-stu-id="337d8-139">x</span></span>|
|[<span data-ttu-id="337d8-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="337d8-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="337d8-141">x</span><span class="sxs-lookup"><span data-stu-id="337d8-141">x</span></span>||
|[<span data-ttu-id="337d8-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="337d8-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="337d8-143">x</span><span class="sxs-lookup"><span data-stu-id="337d8-143">x</span></span>||<span data-ttu-id="337d8-144">x</span><span class="sxs-lookup"><span data-stu-id="337d8-144">x</span></span>|
|[<span data-ttu-id="337d8-145">Rule</span><span class="sxs-lookup"><span data-stu-id="337d8-145">Rule</span></span>](rule.md)||<span data-ttu-id="337d8-146">x</span><span class="sxs-lookup"><span data-stu-id="337d8-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="337d8-147">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="337d8-147">Can contain</span></span>

|<span data-ttu-id="337d8-148">**Élément**</span><span class="sxs-lookup"><span data-stu-id="337d8-148">**Element**</span></span>|<span data-ttu-id="337d8-149">**Contenu**</span><span class="sxs-lookup"><span data-stu-id="337d8-149">**Content**</span></span>|<span data-ttu-id="337d8-150">**Messagerie**</span><span class="sxs-lookup"><span data-stu-id="337d8-150">**Mail**</span></span>|<span data-ttu-id="337d8-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="337d8-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="337d8-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="337d8-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="337d8-153">x</span><span class="sxs-lookup"><span data-stu-id="337d8-153">x</span></span>|<span data-ttu-id="337d8-154">x</span><span class="sxs-lookup"><span data-stu-id="337d8-154">x</span></span>|<span data-ttu-id="337d8-155">x</span><span class="sxs-lookup"><span data-stu-id="337d8-155">x</span></span>|
|[<span data-ttu-id="337d8-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="337d8-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="337d8-157">x</span><span class="sxs-lookup"><span data-stu-id="337d8-157">x</span></span>|<span data-ttu-id="337d8-158">x</span><span class="sxs-lookup"><span data-stu-id="337d8-158">x</span></span>|<span data-ttu-id="337d8-159">x</span><span class="sxs-lookup"><span data-stu-id="337d8-159">x</span></span>|
|[<span data-ttu-id="337d8-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="337d8-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="337d8-161">x</span><span class="sxs-lookup"><span data-stu-id="337d8-161">x</span></span>|<span data-ttu-id="337d8-162">x</span><span class="sxs-lookup"><span data-stu-id="337d8-162">x</span></span>|<span data-ttu-id="337d8-163">x</span><span class="sxs-lookup"><span data-stu-id="337d8-163">x</span></span>|
|[<span data-ttu-id="337d8-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="337d8-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="337d8-165">x</span><span class="sxs-lookup"><span data-stu-id="337d8-165">x</span></span>|<span data-ttu-id="337d8-166">x</span><span class="sxs-lookup"><span data-stu-id="337d8-166">x</span></span>|<span data-ttu-id="337d8-167">x</span><span class="sxs-lookup"><span data-stu-id="337d8-167">x</span></span>|
|[<span data-ttu-id="337d8-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="337d8-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="337d8-169">x</span><span class="sxs-lookup"><span data-stu-id="337d8-169">x</span></span>|<span data-ttu-id="337d8-170">x</span><span class="sxs-lookup"><span data-stu-id="337d8-170">x</span></span>|<span data-ttu-id="337d8-171">x</span><span class="sxs-lookup"><span data-stu-id="337d8-171">x</span></span>|
|[<span data-ttu-id="337d8-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="337d8-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="337d8-173">x</span><span class="sxs-lookup"><span data-stu-id="337d8-173">x</span></span>|<span data-ttu-id="337d8-174">x</span><span class="sxs-lookup"><span data-stu-id="337d8-174">x</span></span>|<span data-ttu-id="337d8-175">x</span><span class="sxs-lookup"><span data-stu-id="337d8-175">x</span></span>|
|[<span data-ttu-id="337d8-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="337d8-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="337d8-177">x</span><span class="sxs-lookup"><span data-stu-id="337d8-177">x</span></span>|<span data-ttu-id="337d8-178">x</span><span class="sxs-lookup"><span data-stu-id="337d8-178">x</span></span>|<span data-ttu-id="337d8-179">x</span><span class="sxs-lookup"><span data-stu-id="337d8-179">x</span></span>|
|[<span data-ttu-id="337d8-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="337d8-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="337d8-181">x</span><span class="sxs-lookup"><span data-stu-id="337d8-181">x</span></span>|||
|[<span data-ttu-id="337d8-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="337d8-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="337d8-183">x</span><span class="sxs-lookup"><span data-stu-id="337d8-183">x</span></span>||
|[<span data-ttu-id="337d8-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="337d8-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="337d8-185">x</span><span class="sxs-lookup"><span data-stu-id="337d8-185">x</span></span>||
|[<span data-ttu-id="337d8-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="337d8-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="337d8-187">x</span><span class="sxs-lookup"><span data-stu-id="337d8-187">x</span></span>|
|[<span data-ttu-id="337d8-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="337d8-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="337d8-189">x</span><span class="sxs-lookup"><span data-stu-id="337d8-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="337d8-190">Attributs</span><span class="sxs-lookup"><span data-stu-id="337d8-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="337d8-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="337d8-191">xmlns</span></span>|<span data-ttu-id="337d8-p101">Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="337d8-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="337d8-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="337d8-194">xmlns:xsi</span></span>|<span data-ttu-id="337d8-p102">Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="337d8-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="337d8-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="337d8-197">xsi:type</span></span>|<span data-ttu-id="337d8-p103">Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="337d8-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
