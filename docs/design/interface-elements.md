---
title: Éléments d’interface utilisateur Office pour les compléments Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 444aca7b75e35ef502075876a7d1324fcdca0603
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32446232"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="b9d04-102">Éléments d’interface utilisateur Office pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="b9d04-102">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="b9d04-p101">Vous pouvez utiliser plusieurs types d’éléments d’interface utilisateur pour étendre l’interface utilisateur d’Office, y compris des commandes de complément et des conteneurs HTML. Ces éléments d’interface utilisateur ressemblent à une extension naturelle d’Office et fonctionnent sur les plateformes. Vous pouvez insérer votre code basé sur le web personnalisé dans l’un de ces éléments.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="b9d04-106">L’image suivante montre les types d’éléments d’interface utilisateur d’Office que vous pouvez créer.</span><span class="sxs-lookup"><span data-stu-id="b9d04-106">The following image shows the types of Office UI elements that you can create.</span></span>

![Image qui affiche des commandes de complément sur le ruban, un volet des tâches et une boîte de dialogue dans un document Office](../images/overview-with-app-interface-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="b9d04-108">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="b9d04-108">Add-in commands</span></span>

<span data-ttu-id="b9d04-p102">Utilisez des [commandes de complément](add-in-commands.md) pour ajouter des points d’entrée vers votre complément au ruban Office. Les commandes démarrent les actions dans votre complément en exécutant du code JavaScript ou en lançant un conteneur HTML. Vous pouvez créer deux types de commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p102">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office ribbon. Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container. You can create two types of add-in commands.</span></span>

|<span data-ttu-id="b9d04-112">**Type de commande**</span><span class="sxs-lookup"><span data-stu-id="b9d04-112">**Command type**</span></span>|<span data-ttu-id="b9d04-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="b9d04-113">**Description**</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="b9d04-114">Onglets, menus et boutons du ruban</span><span class="sxs-lookup"><span data-stu-id="b9d04-114">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="b9d04-p103">Permet d’ajouter des boutons personnalisés, des menus (déroulants) ou des onglets au ruban par défaut dans Office. Utilisez les boutons et menus pour déclencher une action dans Office. Utilisez les onglets pour regrouper et organiser des boutons et menus.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="b9d04-118">Menus contextuels</span><span class="sxs-lookup"><span data-stu-id="b9d04-118">Context menus</span></span>| <span data-ttu-id="b9d04-p104">Permet de développer le menu contextuel par défaut. Les menus contextuels s’affichent lorsque les utilisateurs cliquent avec le bouton droit de la souris sur du texte dans un document Office ou un tableau dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>| 

## <a name="html-containers"></a><span data-ttu-id="b9d04-121">Conteneurs HTML</span><span class="sxs-lookup"><span data-stu-id="b9d04-121">HTML containers</span></span>

<span data-ttu-id="b9d04-p105">Utilisez les conteneurs HTML pour intégrer du code de l’interface utilisateur basé sur HTML dans les clients Office. Ces pages web peuvent ensuite référencer l’API JavaScript Office pour interagir avec du contenu dans le document. Vous pouvez créer trois types de conteneurs HTML.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="b9d04-125">**Conteneur HTML**</span><span class="sxs-lookup"><span data-stu-id="b9d04-125">**HTML container**</span></span>|<span data-ttu-id="b9d04-126">**Description**</span><span class="sxs-lookup"><span data-stu-id="b9d04-126">**Description**</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="b9d04-127">Volets des tâches</span><span class="sxs-lookup"><span data-stu-id="b9d04-127">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="b9d04-p106">Permet d’afficher l’interface utilisateur personnalisée dans le volet droit du document Office. Utilisez les volets des tâches pour permettre aux utilisateurs d’interagir côte à côte avec votre complément et le document Office.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="b9d04-130">Compléments de contenu</span><span class="sxs-lookup"><span data-stu-id="b9d04-130">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="b9d04-p107">Permet d’afficher l’interface utilisateur personnalisée incorporée dans les documents Office. Utilisez les compléments de contenu pour permettre aux utilisateurs d’interagir avec votre complément directement dans le document Office. Par exemple, vous pouvez afficher du contenu externe tel que des vidéos ou des visualisations de données provenant d’autres sources.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="b9d04-134">Boîtes de dialogue</span><span class="sxs-lookup"><span data-stu-id="b9d04-134">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="b9d04-p108">Permet d’afficher l’interface utilisateur personnalisée dans une boîte de dialogue superposée sur le document Office. Utilisez une boîte de dialogue pour les interactions qui nécessitent de l’attention et plus de valeur et ne nécessitent pas une interaction côte-à-côte avec le document.</span><span class="sxs-lookup"><span data-stu-id="b9d04-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="b9d04-137">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b9d04-137">See also</span></span>

- [<span data-ttu-id="b9d04-138">Commandes de complément pour Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b9d04-138">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="b9d04-139">Volets des tâches</span><span class="sxs-lookup"><span data-stu-id="b9d04-139">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="b9d04-140">Compléments de contenu</span><span class="sxs-lookup"><span data-stu-id="b9d04-140">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="b9d04-141">Boîtes de dialogue</span><span class="sxs-lookup"><span data-stu-id="b9d04-141">Dialog boxes</span></span>](dialog-boxes.md)
