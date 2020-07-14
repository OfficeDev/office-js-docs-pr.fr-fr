---
title: Concepts basiques pour les commandes de complément
description: Découvrez l'ajout de boutons et d'éléments de menu personnalisés au ruban dans Office dans le cadre d’un complément Office.
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093874"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="b9039-103">Commandes de complément pour Excel, PowerPoint et Word</span><span class="sxs-lookup"><span data-stu-id="b9039-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="b9039-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span><span class="sxs-lookup"><span data-stu-id="b9039-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span></span> <span data-ttu-id="b9039-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span><span class="sxs-lookup"><span data-stu-id="b9039-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span></span> <span data-ttu-id="b9039-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span><span class="sxs-lookup"><span data-stu-id="b9039-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> <span data-ttu-id="b9039-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span><span class="sxs-lookup"><span data-stu-id="b9039-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="b9039-108">Pour une vue d'ensemble du reportage, voir la vidéo [Ruban de l'application commandes complémentaires au sein du Bureau](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="b9039-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="b9039-109">SharePoint catalogs do not support add-in commands.</span><span class="sxs-lookup"><span data-stu-id="b9039-109">SharePoint catalogs do not support add-in commands.</span></span> <span data-ttu-id="b9039-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span><span class="sxs-lookup"><span data-stu-id="b9039-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9039-111">Les commandes de complément sont actuellement prises en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9039-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="b9039-112">Pour plus d’informations, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="b9039-113">*Figure 1. Complément incluant des commandes en cours d’exécution dans Excel (version de bureau)*</span><span class="sxs-lookup"><span data-stu-id="b9039-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Capture d’écran d’une commande de complément dans Excel](../images/add-in-commands-1.png)

<span data-ttu-id="b9039-115">*Figure 2. Complément incluant des commandes en cours d’exécution dans Excel sur le web*</span><span class="sxs-lookup"><span data-stu-id="b9039-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Capture d’écran d’une commande de complément dans Excel sur le web](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="b9039-117">Fonctionnalités de commande</span><span class="sxs-lookup"><span data-stu-id="b9039-117">Command capabilities</span></span>

<span data-ttu-id="b9039-118">Les fonctionnalités de commande suivantes sont actuellement prises en charge.</span><span class="sxs-lookup"><span data-stu-id="b9039-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="b9039-119">Les compléments de contenu ne prennent actuellement pas en charge les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="b9039-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="b9039-120">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="b9039-120">Extension points</span></span>

- <span data-ttu-id="b9039-121">Onglets de ruban - Permet d’étendre les onglets prédéfinis ou de créer un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="b9039-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="b9039-122">Menus contextuels - Permet d’étendre les menus contextuels sélectionnés.</span><span class="sxs-lookup"><span data-stu-id="b9039-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="b9039-123">Types de contrôles</span><span class="sxs-lookup"><span data-stu-id="b9039-123">Control types</span></span>

- <span data-ttu-id="b9039-124">Boutons simples - Permettent de déclencher des actions spécifiques.</span><span class="sxs-lookup"><span data-stu-id="b9039-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="b9039-125">Menus - Menu déroulant simple avec des boutons qui déclenchent des actions.</span><span class="sxs-lookup"><span data-stu-id="b9039-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="b9039-126">Actions</span><span class="sxs-lookup"><span data-stu-id="b9039-126">Actions</span></span>

- <span data-ttu-id="b9039-127">ShowTaskpane - Affiche un ou plusieurs volets où sont chargées des pages HTML personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b9039-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="b9039-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span><span class="sxs-lookup"><span data-stu-id="b9039-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span></span> <span data-ttu-id="b9039-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span><span class="sxs-lookup"><span data-stu-id="b9039-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="b9039-130">État Activé ou Désactivé par défaut (préversion)</span><span class="sxs-lookup"><span data-stu-id="b9039-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="b9039-131">Vous pouvez spécifier si la commande est activée ou désactivée lors du lancement de votre complément et modifier le paramètre par programme.</span><span class="sxs-lookup"><span data-stu-id="b9039-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="b9039-132">Cette fonctionnalité est en préversion et n’est pas prise en charge dans tous les hôtes ou scénarios.</span><span class="sxs-lookup"><span data-stu-id="b9039-132">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="b9039-133">Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="b9039-134">Plateformes prises en charge</span><span class="sxs-lookup"><span data-stu-id="b9039-134">Supported platforms</span></span>

<span data-ttu-id="b9039-135">Les commandes de complément sont actuellement prises en charge sur les plateformes suivantes.</span><span class="sxs-lookup"><span data-stu-id="b9039-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="b9039-136">Office on Windows (build 16.0.6769+, connecté à l'abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b9039-136">Office on Windows (build 16.0.6769+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="b9039-137">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="b9039-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="b9039-138">Office sur Mac (build 15.33+, connecté à l'abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b9039-138">Office on Mac (build 15.33+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="b9039-139">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="b9039-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="b9039-140">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b9039-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="b9039-141">Pour plus d’informations sur la prise en charge dans Outlook, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="b9039-142">Débogage</span><span class="sxs-lookup"><span data-stu-id="b9039-142">Debugging</span></span>

<span data-ttu-id="b9039-143">Pour déboguer une commande de complément, vous devez l’exécuter dans Office sur le web.</span><span class="sxs-lookup"><span data-stu-id="b9039-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="b9039-144">Pour plus de détails, voir [Débogage de compléments dans Office sur le web](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="b9039-145">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="b9039-145">Best practices</span></span>

<span data-ttu-id="b9039-146">Appliquez les meilleures pratiques suivantes lorsque vous développez des commandes de complément :</span><span class="sxs-lookup"><span data-stu-id="b9039-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="b9039-147">Use commands to represent a specific action with a clear and specific outcome for users.</span><span class="sxs-lookup"><span data-stu-id="b9039-147">Use commands to represent a specific action with a clear and specific outcome for users.</span></span> <span data-ttu-id="b9039-148">Do not combine multiple actions in a single button.</span><span class="sxs-lookup"><span data-stu-id="b9039-148">Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="b9039-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span><span class="sxs-lookup"><span data-stu-id="b9039-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span></span> <span data-ttu-id="b9039-150">Minimize the number of steps an action takes to complete.</span><span class="sxs-lookup"><span data-stu-id="b9039-150">Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="b9039-151">Pour le placement de vos commandes dans le ruban d'application de l'Office :</span><span class="sxs-lookup"><span data-stu-id="b9039-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="b9039-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span><span class="sxs-lookup"><span data-stu-id="b9039-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span></span> <span data-ttu-id="b9039-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span><span class="sxs-lookup"><span data-stu-id="b9039-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span></span> <span data-ttu-id="b9039-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="b9039-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span><span class="sxs-lookup"><span data-stu-id="b9039-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span></span> <span data-ttu-id="b9039-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span><span class="sxs-lookup"><span data-stu-id="b9039-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="b9039-157">Placez des commandes dans un onglet personnalisé si vous avez plus de six commandes de niveau supérieur.</span><span class="sxs-lookup"><span data-stu-id="b9039-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="b9039-158">Name your group to match the name of your add-in.</span><span class="sxs-lookup"><span data-stu-id="b9039-158">Name your group to match the name of your add-in.</span></span> <span data-ttu-id="b9039-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span><span class="sxs-lookup"><span data-stu-id="b9039-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="b9039-160">N’ajoutez pas de boutons superflus pour augmenter la valeur de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b9039-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="b9039-161">Les compléments qui occupent trop d’espace peuvent ne pas obtenir la [validation d’AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="b9039-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="b9039-162">Pour toutes les icônes, suivez les [règles de conception d’icône](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="b9039-163">Proposez une version de complément qui fonctionne aussi sur les hôtes qui ne prennent pas en charge les commandes.</span><span class="sxs-lookup"><span data-stu-id="b9039-163">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="b9039-164">Un seul manifeste de complément peut fonctionner sur les hôtes tenant compte ou non des commandes (par exemple, un volet Office dans le second cas).</span><span class="sxs-lookup"><span data-stu-id="b9039-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="b9039-165">*Figure 3. Complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016*</span><span class="sxs-lookup"><span data-stu-id="b9039-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Capture d’écran illustrant un complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="b9039-167">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="b9039-167">Next steps</span></span>

<span data-ttu-id="b9039-168">La meilleure façon de commencer à utiliser des commandes de complément consiste à consulter des [exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) sur GitHub.</span><span class="sxs-lookup"><span data-stu-id="b9039-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="b9039-169">Pour plus d’informations sur la spécification des commandes de complément dans votre manifeste, reportez-vous à l’article expliquant comment [créer des commandes de complément dans votre manifeste](../develop/create-addin-commands.md) et au contenu de référence sur [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="b9039-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
