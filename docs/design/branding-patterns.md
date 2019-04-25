---
title: Instructions de conception de modèles de personnalisation pour les compléments Office
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 6de9962f82a4d07f94ca34cff5ccc3622f80c5d3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32446998"
---
# <a name="branding-patterns"></a><span data-ttu-id="9fccd-102">Modèles de personnalisation</span><span class="sxs-lookup"><span data-stu-id="9fccd-102">Branding patterns</span></span>

<span data-ttu-id="9fccd-103">Ces modèles assurent la visibilité de la marque et un contexte à vos compléments utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9fccd-103">These patterns provide brand visibilty and context to your add-in users.</span></span> 

## <a name="best-practices"></a><span data-ttu-id="9fccd-104">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="9fccd-104">Best practices</span></span>

|<span data-ttu-id="9fccd-105">À faire</span><span class="sxs-lookup"><span data-stu-id="9fccd-105">Do</span></span> |<span data-ttu-id="9fccd-106">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="9fccd-106">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="9fccd-107">Utilisez des composants d’interface utilisateur familiers en même temps que des caractéristiques de votre marque, comme par exemple une typographie et des couleurs typiques.</span><span class="sxs-lookup"><span data-stu-id="9fccd-107">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="9fccd-108">N’inventez pas des nouveaux composants d’interface utilisateur qui s’opposent aux éléments d’interface utilisateur établis pour Office.</span><span class="sxs-lookup"><span data-stu-id="9fccd-108">Don't invent new UI components that contradict established Office UI.</span></span> | 
| <span data-ttu-id="9fccd-109">Placez la personnalisation de marque pour le complément dans une barre de marque en pied de page en bas de votre interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9fccd-109">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="9fccd-110">Ne répétez pas le nom du volet Office dans une barre de marque immédiatement adjacentes dans la partie supérieure de votre interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9fccd-110">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="9fccd-111">Utilisez les éléments de marque avec parcimonie.</span><span class="sxs-lookup"><span data-stu-id="9fccd-111">Use brand elements sparingly.</span></span> <span data-ttu-id="9fccd-112">Intégrez votre solution à Office pour qu’elle soit complémentaire.</span><span class="sxs-lookup"><span data-stu-id="9fccd-112">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="9fccd-113">N’insérez pas trop d’éléments de personnalisation dans l’interface utilisateur Office, cela risque de détourner l’attention des clients et de les rendre confus.</span><span class="sxs-lookup"><span data-stu-id="9fccd-113">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="9fccd-114">Assurez que votre solution soit facilement reconnaissable et assurez la continuité de vos écrans avec des éléments visuels cohérentes.</span><span class="sxs-lookup"><span data-stu-id="9fccd-114">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="9fccd-115">Ne masquez pas votre solution avec des éléments visuels inconnus et appliqués de manière incohérente.</span><span class="sxs-lookup"><span data-stu-id="9fccd-115">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="9fccd-116">Créez une connexion avec un service ou une entreprise parent pour vous assurer que les clients connaissent et apprécient votre solution.</span><span class="sxs-lookup"><span data-stu-id="9fccd-116">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="9fccd-117">Ne forcez pas les clients à apprendre un nouveau concept de marque s’il existe déjà une relation utile et compréhensible qui peut être utilisée pour créer la confiance et ajouter de la valeur.</span><span class="sxs-lookup"><span data-stu-id="9fccd-117">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |


<span data-ttu-id="9fccd-118">Appliquer les modèles et les composants suivants le cas échéant pour permettre aux utilisateurs de comprendre et utiliser toute l’utilité de votre complément.</span><span class="sxs-lookup"><span data-stu-id="9fccd-118">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>


## <a name="brand-bar"></a><span data-ttu-id="9fccd-119">Barre de marque</span><span class="sxs-lookup"><span data-stu-id="9fccd-119">Brand Bar</span></span>

<span data-ttu-id="9fccd-120">La barre de marque est un espace dans le pied de page où vous pouvez inclure le nom de la marque et le logo.</span><span class="sxs-lookup"><span data-stu-id="9fccd-120">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="9fccd-121">Elle sert également de lien vers le site Web de votre marque et d’emplacement d’accès facultatif.</span><span class="sxs-lookup"><span data-stu-id="9fccd-121">It also serves as a link to your brand's website and an optional access location.</span></span>

![Barre de marque - spécifications pour le volet Office du bureau](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="9fccd-123">Écran de démarrage</span><span class="sxs-lookup"><span data-stu-id="9fccd-123">Splash Screen</span></span>

<span data-ttu-id="9fccd-124">Utilisez cet écran pour afficher votre personnalisation pendant que le complément est en cours de chargement ou lors de la transition entre les différents états de l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9fccd-124">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Écran de démarrage de la marque - spécifications pour le volet Office du bureau](../images/add-in-splash-screen.png)
