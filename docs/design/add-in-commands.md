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
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="3bb68-103">Commandes de complément pour Excel, PowerPoint et Word</span><span class="sxs-lookup"><span data-stu-id="3bb68-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="3bb68-p101">Les commandes de complément sont des éléments d’interface utilisateur qui étendent l’interface utilisateur d’Office et lancent des actions dans votre complément. Vous pouvez les utiliser pour ajouter un bouton sur le ruban ou un élément dans le menu contextuel. Lorsque les utilisateurs sélectionnent une commande de complément, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage d’une page du complément dans le volet Office. Les commandes de complément aident les utilisateurs à trouver et utiliser votre complément, ce qui favorise l’adoption et la réutilisation de votre complément, et améliore la fidélisation des clients.</span><span class="sxs-lookup"><span data-stu-id="3bb68-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="3bb68-108">Pour une vue d'ensemble du reportage, voir la vidéo [Ruban de l'application commandes complémentaires au sein du Bureau](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="3bb68-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="3bb68-p102">Les catalogues SharePoint n’acceptent pas les commandes de complément. Vous pouvez déployer des commandes de complément via le [déploiement centralisé](../publish/centralized-deployment.md) ou [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), ou utiliser le [chargement de version test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) pour déployer votre commande de complément à des fins de test.</span><span class="sxs-lookup"><span data-stu-id="3bb68-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3bb68-111">Les commandes de complément sont actuellement prises en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="3bb68-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="3bb68-112">Pour plus d’informations, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="3bb68-113">*Figure 1. Complément incluant des commandes en cours d’exécution dans Excel (version de bureau)*</span><span class="sxs-lookup"><span data-stu-id="3bb68-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Capture d’écran d’une commande de complément dans Excel](../images/add-in-commands-1.png)

<span data-ttu-id="3bb68-115">*Figure 2. Complément incluant des commandes en cours d’exécution dans Excel sur le web*</span><span class="sxs-lookup"><span data-stu-id="3bb68-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Capture d’écran d’une commande de complément dans Excel sur le web](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="3bb68-117">Fonctionnalités de commande</span><span class="sxs-lookup"><span data-stu-id="3bb68-117">Command capabilities</span></span>

<span data-ttu-id="3bb68-118">Les fonctionnalités de commande suivantes sont actuellement prises en charge.</span><span class="sxs-lookup"><span data-stu-id="3bb68-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="3bb68-119">Les compléments de contenu ne prennent actuellement pas en charge les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="3bb68-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="3bb68-120">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="3bb68-120">Extension points</span></span>

- <span data-ttu-id="3bb68-121">Onglets de ruban - Permet d’étendre les onglets prédéfinis ou de créer un onglet personnalisé.</span><span class="sxs-lookup"><span data-stu-id="3bb68-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="3bb68-122">Menus contextuels - Permet d’étendre les menus contextuels sélectionnés.</span><span class="sxs-lookup"><span data-stu-id="3bb68-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="3bb68-123">Types de contrôles</span><span class="sxs-lookup"><span data-stu-id="3bb68-123">Control types</span></span>

- <span data-ttu-id="3bb68-124">Boutons simples - Permettent de déclencher des actions spécifiques.</span><span class="sxs-lookup"><span data-stu-id="3bb68-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="3bb68-125">Menus - Menu déroulant simple avec des boutons qui déclenchent des actions.</span><span class="sxs-lookup"><span data-stu-id="3bb68-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="3bb68-126">Actions</span><span class="sxs-lookup"><span data-stu-id="3bb68-126">Actions</span></span>

- <span data-ttu-id="3bb68-127">ShowTaskpane - Affiche un ou plusieurs volets où sont chargées des pages HTML personnalisées.</span><span class="sxs-lookup"><span data-stu-id="3bb68-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="3bb68-p104">ExecuteFunction - Charge une page HTML invisible, puis y exécute une fonction JavaScript. Pour afficher l’interface utilisateur au sein de votre fonction (par exemple, erreurs, avancement, entrées supplémentaires), vous pouvez utiliser l’API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="3bb68-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="3bb68-130">État Activé ou Désactivé par défaut (préversion)</span><span class="sxs-lookup"><span data-stu-id="3bb68-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="3bb68-131">Vous pouvez spécifier si la commande est activée ou désactivée lors du lancement de votre complément et modifier le paramètre par programme.</span><span class="sxs-lookup"><span data-stu-id="3bb68-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="3bb68-132">Cette fonctionnalité est en préversion et n’est pas prise en charge dans tous les hôtes ou scénarios.</span><span class="sxs-lookup"><span data-stu-id="3bb68-132">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="3bb68-133">Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="3bb68-134">Plateformes prises en charge</span><span class="sxs-lookup"><span data-stu-id="3bb68-134">Supported platforms</span></span>

<span data-ttu-id="3bb68-135">Les commandes de complément sont actuellement prises en charge sur les plateformes suivantes.</span><span class="sxs-lookup"><span data-stu-id="3bb68-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="3bb68-136">Office on Windows (build 16.0.6769+, connecté à l'abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3bb68-136">Office on Windows (build 16.0.6769+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3bb68-137">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="3bb68-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="3bb68-138">Office sur Mac (build 15.33+, connecté à l'abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3bb68-138">Office on Mac (build 15.33+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3bb68-139">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="3bb68-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="3bb68-140">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="3bb68-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="3bb68-141">Pour plus d’informations sur la prise en charge dans Outlook, voir [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="3bb68-142">Débogage</span><span class="sxs-lookup"><span data-stu-id="3bb68-142">Debugging</span></span>

<span data-ttu-id="3bb68-143">Pour déboguer une commande de complément, vous devez l’exécuter dans Office sur le web.</span><span class="sxs-lookup"><span data-stu-id="3bb68-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="3bb68-144">Pour plus de détails, voir [Débogage de compléments dans Office sur le web](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="3bb68-145">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="3bb68-145">Best practices</span></span>

<span data-ttu-id="3bb68-146">Appliquez les meilleures pratiques suivantes lorsque vous développez des commandes de complément :</span><span class="sxs-lookup"><span data-stu-id="3bb68-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="3bb68-p107">Utilisez les commandes pour représenter une action spécifique avec un résultat clair et précis pour les utilisateurs. Ne combinez pas plusieurs actions dans un seul bouton.</span><span class="sxs-lookup"><span data-stu-id="3bb68-p107">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="3bb68-p108">Proposez des actions détaillées permettant de réaliser plus efficacement des tâches courantes dans votre complément. Réduisez le nombre d’étapes nécessaires à la réalisation d’une action.</span><span class="sxs-lookup"><span data-stu-id="3bb68-p108">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="3bb68-151">Pour le placement de vos commandes dans le ruban d'application de l'Office :</span><span class="sxs-lookup"><span data-stu-id="3bb68-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="3bb68-p109">Placez les commandes sur un onglet existant (Insertion, Révision, etc.) si la fonctionnalité ajoutée lui correspond. Par exemple, si votre complément permet aux utilisateurs d’insérer un élément multimédia, ajoutez un groupe à l’onglet Insertion. Notez que l’ensemble des onglets ne sont pas nécessairement disponibles dans toutes les versions d’Office. Pour plus d’informations, voir le [manifeste XML de compléments Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-p109">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="3bb68-p110">Placez les commandes sous l’onglet Accueil si la fonctionnalité ne correspond à aucun autre onglet, et si vous avez moins de six commandes de niveau supérieur. Vous pouvez également ajouter des commandes à l’onglet Accueil si votre complément doit fonctionner sur toutes les versions d’Office (par exemple, Office sur le web ou le bureau) et si un onglet n’est pas disponible dans toutes les versions (par exemple, si l’onglet Création n’existe pas dans Office sur le web).</span><span class="sxs-lookup"><span data-stu-id="3bb68-p110">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="3bb68-157">Placez des commandes dans un onglet personnalisé si vous avez plus de six commandes de niveau supérieur.</span><span class="sxs-lookup"><span data-stu-id="3bb68-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="3bb68-p111">Nommez votre groupe en fonction du nom de votre complément. Si vous avez plusieurs groupes, nommez chaque groupe en fonction de la fonctionnalité offerte par les commandes de ce groupe.</span><span class="sxs-lookup"><span data-stu-id="3bb68-p111">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="3bb68-160">N’ajoutez pas de boutons superflus pour augmenter la valeur de votre complément.</span><span class="sxs-lookup"><span data-stu-id="3bb68-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3bb68-161">Les compléments qui occupent trop d’espace peuvent ne pas obtenir la [validation d’AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="3bb68-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="3bb68-162">Pour toutes les icônes, suivez les [règles de conception d’icône](add-in-icons.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="3bb68-163">Proposez une version de complément qui fonctionne aussi sur les hôtes qui ne prennent pas en charge les commandes.</span><span class="sxs-lookup"><span data-stu-id="3bb68-163">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="3bb68-164">Un seul manifeste de complément peut fonctionner sur les hôtes tenant compte ou non des commandes (par exemple, un volet Office dans le second cas).</span><span class="sxs-lookup"><span data-stu-id="3bb68-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="3bb68-165">*Figure 3. Complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016*</span><span class="sxs-lookup"><span data-stu-id="3bb68-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Capture d’écran illustrant un complément du volet Office dans Office 2013 et le même complément utilisant des commandes de complément dans Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="3bb68-167">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="3bb68-167">Next steps</span></span>

<span data-ttu-id="3bb68-168">La meilleure façon de commencer à utiliser des commandes de complément consiste à consulter des [exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) sur GitHub.</span><span class="sxs-lookup"><span data-stu-id="3bb68-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="3bb68-169">Pour plus d’informations sur la spécification des commandes de complément dans votre manifeste, reportez-vous à l’article expliquant comment [créer des commandes de complément dans votre manifeste](../develop/create-addin-commands.md) et au contenu de référence sur [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="3bb68-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
