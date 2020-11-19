---
title: Modèles de navigation pour les compléments Office
description: Découvrez les meilleures pratiques pour l’utilisation des barres de commandes, des barres d’onglets et des boutons de retour, pour concevoir la navigation d’un complément Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132031"
---
# <a name="navigation-patterns"></a><span data-ttu-id="e101a-103">Modèles de navigation</span><span class="sxs-lookup"><span data-stu-id="e101a-103">Navigation patterns</span></span>

<span data-ttu-id="e101a-104">Les principales fonctionnalités d’un complément sont accessibles via les types de commande spécifique et la zone de l’écran limitée.</span><span class="sxs-lookup"><span data-stu-id="e101a-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="e101a-105">Il est important que la navigation soit intuitive, fournisse du contexte et permette à l’utilisateur de se déplacer facilement au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="e101a-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="e101a-106">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="e101a-106">Best practices</span></span>

| <span data-ttu-id="e101a-107">À faire</span><span class="sxs-lookup"><span data-stu-id="e101a-107">Do</span></span>    | <span data-ttu-id="e101a-108">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="e101a-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="e101a-109">Vérifiez que l’utilisateur dispose d’une option de navigation clairement visible.</span><span class="sxs-lookup"><span data-stu-id="e101a-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="e101a-110">Ne compliquez pas le processus de navigation en utilisant des éléments d’interface utilisateur non standard.</span><span class="sxs-lookup"><span data-stu-id="e101a-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="e101a-111">Servez-vous des composants suivants le cas échéant pour permettre aux utilisateurs de parcourir le complément.</span><span class="sxs-lookup"><span data-stu-id="e101a-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="e101a-112">N’ajoutez pas de difficultés qui empêcherait l’utilisateur de savoir où il se trouve ou de comprendre le contexte au sein du complément</span><span class="sxs-lookup"><span data-stu-id="e101a-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>

## <a name="command-bar"></a><span data-ttu-id="e101a-113">Barre de commandes</span><span class="sxs-lookup"><span data-stu-id="e101a-113">Command Bar</span></span>

<span data-ttu-id="e101a-114">La barre de commandes est une surface dans le volet Office qui héberge des commandes qui fonctionnent sur le contenu de la fenêtre, du panneau ou de la région parent qu’elle contient.</span><span class="sxs-lookup"><span data-stu-id="e101a-114">The CommandBar is a surface within the task pane that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="e101a-115">Exemples de fonctionnalités facultatives : point d’accès au menu « hamburger », recherche et commandes sur le côté.</span><span class="sxs-lookup"><span data-stu-id="e101a-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Illustration d’une barre de commandes dans un volet Office d’une application de bureau Office.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a><span data-ttu-id="e101a-118">Barre d’onglets</span><span class="sxs-lookup"><span data-stu-id="e101a-118">Tab Bar</span></span>

<span data-ttu-id="e101a-119">La barre d’onglets affiche la navigation à l’aide de boutons avec du texte et des icônes verticalement empilés.</span><span class="sxs-lookup"><span data-stu-id="e101a-119">The tab bar shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="e101a-120">Utiliser la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.</span><span class="sxs-lookup"><span data-stu-id="e101a-120">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Illustration d’une barre d’onglets dans le volet Office d’une application de bureau Office.](../images/add-in-tab-bar.png)

## <a name="back-button"></a><span data-ttu-id="e101a-123">Bouton Précédent</span><span class="sxs-lookup"><span data-stu-id="e101a-123">Back Button</span></span>

<span data-ttu-id="e101a-124">Le bouton précédent permet aux utilisateurs de récupérer à partir d’une action de navigation d’exploration.</span><span class="sxs-lookup"><span data-stu-id="e101a-124">The back button allows users to recover from a drill-down navigational action.</span></span> <span data-ttu-id="e101a-125">Ce modèle permet de vous assurer que les utilisateurs suivent une série d’étapes ordonnées.</span><span class="sxs-lookup"><span data-stu-id="e101a-125">This pattern helps ensure users follow an ordered series of steps.</span></span>

![Illustration illustrant un bouton retour dans le volet Office d’une application de bureau Office.](../images/add-in-back-button.png)
