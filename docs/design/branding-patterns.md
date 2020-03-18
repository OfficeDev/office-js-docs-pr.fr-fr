---
title: Instructions de conception de modèles de personnalisation pour les compléments Office
description: Découvrez comment personnaliser votre complément Office tout en restant compatible avec la conception visuelle d’Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: d2f492f5f1654c6bd6448db4c2d1707c26b42af9
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717249"
---
# <a name="branding-patterns"></a><span data-ttu-id="d1cee-103">Modèles de personnalisation</span><span class="sxs-lookup"><span data-stu-id="d1cee-103">Branding patterns</span></span>

<span data-ttu-id="d1cee-104">Ces modèles assurent la visibilité de la marque et un contexte à vos compléments utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d1cee-104">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="d1cee-105">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="d1cee-105">Best practices</span></span>

|<span data-ttu-id="d1cee-106">À faire</span><span class="sxs-lookup"><span data-stu-id="d1cee-106">Do</span></span> |<span data-ttu-id="d1cee-107">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="d1cee-107">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="d1cee-108">Utilisez des composants d’interface utilisateur familiers en même temps que des caractéristiques de votre marque, comme par exemple une typographie et des couleurs typiques.</span><span class="sxs-lookup"><span data-stu-id="d1cee-108">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="d1cee-109">N’inventez pas des nouveaux composants d’interface utilisateur qui s’opposent aux éléments d’interface utilisateur établis pour Office.</span><span class="sxs-lookup"><span data-stu-id="d1cee-109">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="d1cee-110">Placez la personnalisation de marque pour le complément dans une barre de marque en pied de page en bas de votre interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d1cee-110">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="d1cee-111">Ne répétez pas le nom du volet Office dans une barre de marque immédiatement adjacentes dans la partie supérieure de votre interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d1cee-111">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="d1cee-112">Utilisez les éléments de marque avec parcimonie.</span><span class="sxs-lookup"><span data-stu-id="d1cee-112">Use brand elements sparingly.</span></span> <span data-ttu-id="d1cee-113">Intégrez votre solution à Office pour qu’elle soit complémentaire.</span><span class="sxs-lookup"><span data-stu-id="d1cee-113">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="d1cee-114">N’insérez pas trop d’éléments de personnalisation dans l’interface utilisateur Office, cela risque de détourner l’attention des clients et de les rendre confus.</span><span class="sxs-lookup"><span data-stu-id="d1cee-114">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="d1cee-115">Assurez que votre solution soit facilement reconnaissable et assurez la continuité de vos écrans avec des éléments visuels cohérentes.</span><span class="sxs-lookup"><span data-stu-id="d1cee-115">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="d1cee-116">Ne masquez pas votre solution avec des éléments visuels inconnus et appliqués de manière incohérente.</span><span class="sxs-lookup"><span data-stu-id="d1cee-116">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="d1cee-117">Créez une connexion avec un service ou une entreprise parent pour vous assurer que les clients connaissent et apprécient votre solution.</span><span class="sxs-lookup"><span data-stu-id="d1cee-117">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="d1cee-118">Ne forcez pas les clients à apprendre un nouveau concept de marque s’il existe déjà une relation utile et compréhensible qui peut être utilisée pour créer la confiance et ajouter de la valeur.</span><span class="sxs-lookup"><span data-stu-id="d1cee-118">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="d1cee-119">Appliquer les modèles et les composants suivants le cas échéant pour permettre aux utilisateurs de comprendre et utiliser toute l’utilité de votre complément.</span><span class="sxs-lookup"><span data-stu-id="d1cee-119">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="d1cee-120">Barre de marque</span><span class="sxs-lookup"><span data-stu-id="d1cee-120">Brand Bar</span></span>

<span data-ttu-id="d1cee-121">La barre de marque est un espace dans le pied de page où vous pouvez inclure le nom de la marque et le logo.</span><span class="sxs-lookup"><span data-stu-id="d1cee-121">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="d1cee-122">Elle sert également de lien vers le site Web de votre marque et d’emplacement d’accès facultatif.</span><span class="sxs-lookup"><span data-stu-id="d1cee-122">It also serves as a link to your brand's website and an optional access location.</span></span>

![Barre de marque - spécifications pour le volet Office du bureau](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="d1cee-124">Écran de démarrage</span><span class="sxs-lookup"><span data-stu-id="d1cee-124">Splash Screen</span></span>

<span data-ttu-id="d1cee-125">Utilisez cet écran pour afficher votre personnalisation pendant que le complément est en cours de chargement ou lors de la transition entre les différents états de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d1cee-125">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Écran de démarrage de la marque - spécifications pour le volet Office du bureau](../images/add-in-splash-screen.png)
