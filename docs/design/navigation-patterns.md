---
title: Modèles de navigation pour les compléments Office
description: ''
ms.date: 06/26/2018
ms.openlocfilehash: b7fee6fad703ce7c8f4c5f8b848d6bf28b239b09
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432969"
---
# <a name="navigation-patterns"></a><span data-ttu-id="7f6b8-102">Modèles de navigation</span><span class="sxs-lookup"><span data-stu-id="7f6b8-102">Navigation patterns</span></span>

<span data-ttu-id="7f6b8-103">Les principales fonctionnalités d’un complément sont accessibles via les types de commande spécifique et la zone de l’écran limitée.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-103">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="7f6b8-104">Il est important que la navigation soit intuitive, fournisse du contexte et permette à l’utilisateur de se déplacer facilement au sein du complément.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-104">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="7f6b8-105">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="7f6b8-105">Best practices</span></span>

| <span data-ttu-id="7f6b8-106">À faire</span><span class="sxs-lookup"><span data-stu-id="7f6b8-106">Do</span></span>    | <span data-ttu-id="7f6b8-107">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="7f6b8-107">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="7f6b8-108">Vérifiez que l’utilisateur dispose d’une option de navigation clairement visible.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-108">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="7f6b8-109">Ne compliquez pas le processus de navigation en utilisant des éléments d’interface utilisateur non standard.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-109">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="7f6b8-110">Servez-vous des composants suivants le cas échéant pour permettre aux utilisateurs de parcourir le complément.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-110">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="7f6b8-111">N’ajoutez pas de difficultés qui empêcherait l’utilisateur de savoir où il se trouve ou de comprendre le contexte au sein du complément</span><span class="sxs-lookup"><span data-stu-id="7f6b8-111">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="7f6b8-112">Barre de commandes</span><span class="sxs-lookup"><span data-stu-id="7f6b8-112">UserForm Command Bar</span></span>

<span data-ttu-id="7f6b8-113">CommandBar est une surface qui héberge les commandes qui fonctionnent sur le contenu de la fenêtre, du panneau de configuration ou de la région parent située au-dessous.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-113">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="7f6b8-114">Exemples de fonctionnalités facultatives : point d’accès au menu « hamburger », recherche et commandes sur le côté.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-114">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Commandes – spécifications pour le volet Office du bureau](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="7f6b8-116">Barre d’onglets</span><span class="sxs-lookup"><span data-stu-id="7f6b8-116">Tab bar</span></span>

<span data-ttu-id="7f6b8-117">Affiche la navigation à l’aide de boutons avec du texte et des icônes empilés verticalement.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-117">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="7f6b8-118">Utilisez la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-118">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Barre d’onglets – spécifications pour le volet Office du bureau](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="7f6b8-120">Bouton Précédent</span><span class="sxs-lookup"><span data-stu-id="7f6b8-120">Back button</span></span>

<span data-ttu-id="7f6b8-121">Le bouton Précédent permet aux utilisateurs de revenir en arrière après avoir navigué dans l’interface.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-121">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="7f6b8-122">Ce modèle permet de vous assurer que les utilisateurs suivent une série d’étapes ordonnées.</span><span class="sxs-lookup"><span data-stu-id="7f6b8-122">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![Bouton Précédent – spécifications pour le volet Office du bureau](../images/add-in-back-button.png)
