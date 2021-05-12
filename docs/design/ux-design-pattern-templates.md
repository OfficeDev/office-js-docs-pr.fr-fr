---
title: Modèles de conception de l’expérience utilisateur pour les compléments Office
description: Obtenez une vue d’ensemble des modèles de conception d’interface utilisateur pour les Office, y compris les modèles de navigation, d’authentification, de première utilisation et de authentification.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 8544b56b85a25d522c95546b42a78fe01a3c2586
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330107"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="8ab9a-103">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="8ab9a-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="8ab9a-104">La conception de l’expérience utilisateur pour les compléments Office doit offrir une expérience attrayante aux utilisateurs d’Office et étendre l’expérience générale Office en s'intégrant parfaitement dans l’interface utilisateur Office par défaut.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="8ab9a-105">Nos modèles d’expérience utilisateur sont composés de composants.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="8ab9a-106">Les composants sont des contrôles qui aident vos clients à interagir avec les éléments de votre logiciel ou service.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="8ab9a-107">Les boutons, la navigation et les menus sont des exemples de composants courants qui ont souvent des comportements et des styles cohérents.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="8ab9a-108">[L’interface utilisateur Fluent React composants](using-office-ui-fabric-react.md) se comportent comme une partie de Office, tout comme les composants neutres de l’infrastructure de [Office UI Fabric JS](fabric-core.md).</span><span class="sxs-lookup"><span data-stu-id="8ab9a-108">[Fluent UI React components](using-office-ui-fabric-react.md) look and behave like a part of Office, as do the framework-neutral components of [Office UI Fabric JS](fabric-core.md).</span></span> <span data-ttu-id="8ab9a-109">Tirez parti de l’un ou l’autre des ensembles de composants à intégrer à Office.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-109">Take advantage of either set of components to integrate with Office.</span></span> <span data-ttu-id="8ab9a-110">Sinon, si votre add-in possède son propre langage de composant existant, vous n’avez pas besoin de l’ignorer.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-110">Alternatively, if your add-in has its own preexisting component language, you don't need to discard it.</span></span> <span data-ttu-id="8ab9a-111">Recherchez les opportunités pour le conserver lors de l’intégration avec Office.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="8ab9a-112">Pensez à remplacer les éléments stylistiques, à supprimer les conflits ou à adopter des styles et des comportements qui éliminent la confusion de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="8ab9a-113">Les modèles fournis sont les meilleures solutions pratiques basées sur des scénarios courants d’utilisation et sur les recherches en expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="8ab9a-114">Ils sont destinés à fournir un point d’entrée rapide pour concevoir et développer des modules, ainsi que des conseils pour trouver un équilibre entre les éléments de marque Microsoft et les vôtres.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own.</span></span> <span data-ttu-id="8ab9a-115">Fournir une expérience utilisateur propre et moderne qui équilibre les éléments de conception du langage de conception de l’interface utilisateur Fluent de Microsoft et l’identité de marque unique du partenaire peut aider à augmenter la rétention et l’adoption par les utilisateurs de votre add-in.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="8ab9a-116">Utiliser les modèles de motif expérience utilisateur pour :</span><span class="sxs-lookup"><span data-stu-id="8ab9a-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="8ab9a-117">Appliquer des solutions à des scénarios client courants.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="8ab9a-118">Appliquer les meilleures pratiques en matière de conception.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-118">Apply design best practices.</span></span>
* <span data-ttu-id="8ab9a-119">Incorporer [des composants et des](https://developer.microsoft.com/fluentui#/get-started) styles d’interface utilisateur Fluent.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-119">Incorporate [Fluent UI](https://developer.microsoft.com/fluentui#/get-started) components and styles.</span></span>
* <span data-ttu-id="8ab9a-120">Créer des compléments qui s’intègrent visuellement à l’interface utilisateur d’Office par défaut.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="8ab9a-121">Imaginer et visualiser l’expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="8ab9a-122">Prise en main</span><span class="sxs-lookup"><span data-stu-id="8ab9a-122">Getting started</span></span>

<span data-ttu-id="8ab9a-123">Les modèles sont organisés par les actions clés ou les expériences qui sont courantes dans un complément.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="8ab9a-124">Les groupes principaux sont :</span><span class="sxs-lookup"><span data-stu-id="8ab9a-124">The main groups are:</span></span>

* [<span data-ttu-id="8ab9a-125">Première exécution</span><span class="sxs-lookup"><span data-stu-id="8ab9a-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="8ab9a-126">Authentification</span><span class="sxs-lookup"><span data-stu-id="8ab9a-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="8ab9a-127">Navigation</span><span class="sxs-lookup"><span data-stu-id="8ab9a-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="8ab9a-128">Conception de personnalisation</span><span class="sxs-lookup"><span data-stu-id="8ab9a-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="8ab9a-129">Étudiez chaque groupe pour apprendre comment concevoir votre complément en utilisant les meilleures pratiques.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab9a-130">Les écrans exemple illustrés dans l’ensemble de cette documentation sont conçus et affichés à une résolution de **1366 x 768**.</span><span class="sxs-lookup"><span data-stu-id="8ab9a-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="8ab9a-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8ab9a-131">See also</span></span>

* [<span data-ttu-id="8ab9a-132">Kits d’outils de conception</span><span class="sxs-lookup"><span data-stu-id="8ab9a-132">Design tool kits</span></span>](design-toolkits.md)
* [<span data-ttu-id="8ab9a-133">Interface utilisateur Fluent</span><span class="sxs-lookup"><span data-stu-id="8ab9a-133">Fluent UI</span></span>](https://developer.microsoft.com/fluentui#)
* [<span data-ttu-id="8ab9a-134">Meilleures pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="8ab9a-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="8ab9a-135">Interface utilisateur Fluent React dans Office de l’interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="8ab9a-135">Fluent UI React in Office Add-ins</span></span>](using-office-ui-fabric-react.md)
