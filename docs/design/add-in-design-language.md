---
title: Langage de création d’un complément Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e0975f8ec5c0706509dbb7d1fb39defc6c21e006
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925086"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="7017a-102">Langage de création d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="7017a-102">Office Add-in design language</span></span>

<span data-ttu-id="7017a-p101">Le langage de création d’Office est un système visuel clair et simple qui garantit la cohérence entre expériences. Il contient un ensemble d’éléments visuels qui définissent les interfaces Office, y compris :</span><span class="sxs-lookup"><span data-stu-id="7017a-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="7017a-105">Police standard</span><span class="sxs-lookup"><span data-stu-id="7017a-105">A standard typeface</span></span>
- <span data-ttu-id="7017a-106">Palette de couleurs courantes</span><span class="sxs-lookup"><span data-stu-id="7017a-106">A common color palette</span></span>
- <span data-ttu-id="7017a-107">Ensemble de tailles typographiques et pondérations</span><span class="sxs-lookup"><span data-stu-id="7017a-107">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="7017a-108">Instructions relatives aux icônes</span><span class="sxs-lookup"><span data-stu-id="7017a-108">Icon guidelines</span></span>
- <span data-ttu-id="7017a-109">Éléments d’icône partagée</span><span class="sxs-lookup"><span data-stu-id="7017a-109">Shared icon assets</span></span>
- <span data-ttu-id="7017a-110">Définitions d’animation</span><span class="sxs-lookup"><span data-stu-id="7017a-110">Animation definitions</span></span>
- <span data-ttu-id="7017a-111">Composants courants</span><span class="sxs-lookup"><span data-stu-id="7017a-111">Common components</span></span>

<span data-ttu-id="7017a-p102">[Office UI Fabric](https://developer.microsoft.com/fabric) est l’infrastructure frontale officielle pour la création avec le langage de création Office. L’utilisation de Fabric est facultative, mais elle est le moyen le plus rapide pour vous assurer que vos compléments sont une extension naturelle d’Office. Profitez de Fabric pour concevoir et créer des compléments qui complètent Office.</span><span class="sxs-lookup"><span data-stu-id="7017a-p102">[Office UI Fabric](https://developer.microsoft.com/fabric) is the official front-end framework for building with the Office design language. Using Fabric is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office. Take advantage of Fabric to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="7017a-p103">De nombreux compléments d’Office sont associés à une marque préexistante. Vous pouvez conserver une marque forte et son langage de composant ou visuel dans votre complément. Recherchez les opportunités pour conserver votre propre langage visuel lors de l’intégration avec Office. Pensez à des moyens de remplacer les couleurs Office, la typographie, les icônes ou d’autres éléments stylistiques par des éléments de votre marque. Pensez à des moyens de suivre des dispositions de complément ou des modèles de conception de l’expérience utilisateur courants tout en insérant des contrôles et des composants que vos clients connaissent.</span><span class="sxs-lookup"><span data-stu-id="7017a-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="7017a-p104">L’insertion d’une interface utilisateur HTML de marque importante à l’intérieur d’Office peut créer des dissonances pour les clients. Trouvez un équilibre qui s’adapte en toute transparence dans Office mais qui s’aligne aussi clairement sur votre marque parent ou de service. Lorsqu’un complément ne s’adapte pas à Office, c’est souvent en raison d’une incompatibilité des éléments stylistiques. Par exemple, la typographie est trop grande et en dehors de la grille, les couleurs sont particulièrement criardes ou contrastées, ou les animations sont superflues et se comportent différemment par rapport à Office. L’apparence et le comportement des contrôles ou des composants dévient trop des normes d’Office.</span><span class="sxs-lookup"><span data-stu-id="7017a-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
