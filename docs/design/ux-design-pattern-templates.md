---
title: Modèles de conception de l’expérience utilisateur pour les compléments Office
description: Obtenez une vue d’ensemble des modèles de conception de l’interface utilisateur pour les compléments Office, y compris les modèles de navigation, d’authentification, de première utilisation et de personnalisation.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d7201cd91dbfd019a7b045a7f63c1c86a74b9142
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608459"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="62d7b-103">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="62d7b-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="62d7b-104">La conception de l’expérience utilisateur pour les compléments Office doit offrir une expérience attrayante aux utilisateurs d’Office et étendre l’expérience générale Office en s'intégrant parfaitement dans l’interface utilisateur Office par défaut.</span><span class="sxs-lookup"><span data-stu-id="62d7b-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="62d7b-105">Nos modèles d’expérience utilisateur sont composés de composants.</span><span class="sxs-lookup"><span data-stu-id="62d7b-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="62d7b-106">Les composants sont des contrôles qui aident vos clients à interagir avec les éléments de votre logiciel ou service.</span><span class="sxs-lookup"><span data-stu-id="62d7b-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="62d7b-107">Les boutons, la navigation et les menus sont des exemples de composants courants qui ont souvent des comportements et des styles cohérents.</span><span class="sxs-lookup"><span data-stu-id="62d7b-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="62d7b-108">Office UI Fabric rend les composants qui ressemblent à une partie d’Office et se comportent comme une partie d’Office.</span><span class="sxs-lookup"><span data-stu-id="62d7b-108">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="62d7b-109">Utilisez Fabric pour une intégration facile avec Office.</span><span class="sxs-lookup"><span data-stu-id="62d7b-109">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="62d7b-110">Si votre complément a son propre langage de composant préexistant, vous n’avez pas besoin de l’abandonner en faveur de Fabric.</span><span class="sxs-lookup"><span data-stu-id="62d7b-110">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="62d7b-111">Recherchez les opportunités pour le conserver lors de l’intégration avec Office.</span><span class="sxs-lookup"><span data-stu-id="62d7b-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="62d7b-112">Pensez à remplacer les éléments stylistiques, à supprimer les conflits ou à adopter des styles et des comportements qui éliminent la confusion de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="62d7b-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="62d7b-113">Les modèles fournis sont les meilleures solutions pratiques basées sur des scénarios courants d’utilisation et sur les recherches en expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="62d7b-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="62d7b-114">Ils sont destinés à fournir une entrée rapide à la conception et au développement de compléments, et fournir des conseils pour obtenir un équilibre entre les éléments Microsoft et les éléments de la marque.</span><span class="sxs-lookup"><span data-stu-id="62d7b-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="62d7b-115">Fournir une expérience utilisateur propre et moderne qui assure un équilibre entre les éléments de conception du langage de conception Microsoft Fabric et l’identité de marque unique du partenaire peut vous aider à augmenter la rétention utilisateur et l’adoption de votre complément.</span><span class="sxs-lookup"><span data-stu-id="62d7b-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="62d7b-116">Utiliser les modèles de motif expérience utilisateur pour :</span><span class="sxs-lookup"><span data-stu-id="62d7b-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="62d7b-117">Appliquer des solutions à des scénarios client courants.</span><span class="sxs-lookup"><span data-stu-id="62d7b-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="62d7b-118">Appliquer les meilleures pratiques en matière de conception.</span><span class="sxs-lookup"><span data-stu-id="62d7b-118">Apply design best practices.</span></span>
* <span data-ttu-id="62d7b-119">Incorporer les composants et styles d’[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started).</span><span class="sxs-lookup"><span data-stu-id="62d7b-119">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="62d7b-120">Créer des compléments qui s’intègrent visuellement à l’interface utilisateur d’Office par défaut.</span><span class="sxs-lookup"><span data-stu-id="62d7b-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="62d7b-121">Imaginer et visualiser l’expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="62d7b-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="62d7b-122">Prise en main</span><span class="sxs-lookup"><span data-stu-id="62d7b-122">Getting started</span></span>

<span data-ttu-id="62d7b-123">Les modèles sont organisés par les actions clés ou les expériences qui sont courantes dans un complément.</span><span class="sxs-lookup"><span data-stu-id="62d7b-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="62d7b-124">Les groupes principaux sont :</span><span class="sxs-lookup"><span data-stu-id="62d7b-124">The main groups are:</span></span>

* [<span data-ttu-id="62d7b-125">Première exécution</span><span class="sxs-lookup"><span data-stu-id="62d7b-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="62d7b-126">Authentification</span><span class="sxs-lookup"><span data-stu-id="62d7b-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="62d7b-127">Navigation</span><span class="sxs-lookup"><span data-stu-id="62d7b-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="62d7b-128">Conception de personnalisation</span><span class="sxs-lookup"><span data-stu-id="62d7b-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="62d7b-129">Étudiez chaque groupe pour apprendre comment concevoir votre complément en utilisant les meilleures pratiques.</span><span class="sxs-lookup"><span data-stu-id="62d7b-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="62d7b-130">Les écrans exemple illustrés dans l’ensemble de cette documentation sont conçus et affichés à une résolution de **1366 x 768**.</span><span class="sxs-lookup"><span data-stu-id="62d7b-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="62d7b-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="62d7b-131">See also</span></span>

* [<span data-ttu-id="62d7b-132">Kits de ressources de conception</span><span class="sxs-lookup"><span data-stu-id="62d7b-132">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="62d7b-133">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="62d7b-133">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="62d7b-134">Meilleures pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="62d7b-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="62d7b-135">Prise en main de Fabric React</span><span class="sxs-lookup"><span data-stu-id="62d7b-135">Get started using Fabric React</span></span>](../design/using-office-ui-fabric-react.md)
