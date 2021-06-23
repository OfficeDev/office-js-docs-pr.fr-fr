---
title: Modèles de première expérience d’utilisation des complément Office
description: Découvrez les meilleures pratiques pour concevoir des expériences de première Office des modules.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: d020a281aca10805ba8fd1176403f3788f6d716c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076342"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="94e6b-103">Modèles de première expérience d’utilisation</span><span class="sxs-lookup"><span data-stu-id="94e6b-103">First-run experience patterns</span></span>

<span data-ttu-id="94e6b-104">Une première expérience d’utilisation (FRE) correspond à l’introduction d’un utilisateur à votre complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="94e6b-105">Une FRE existe quand un utilisateur ouvre un complément pour la première fois et lui fournit des informations sur les fonctions, les fonctionnalités et/ou les avantages du complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="94e6b-106">Cette expérience vous permet de modeler la première impression qu’un utilisateur va avoir d’un complément. Elle peut grandement influencer la probabilité qu’il y revienne et continue à utiliser votre complément...</span><span class="sxs-lookup"><span data-stu-id="94e6b-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="94e6b-107">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="94e6b-107">Best practices</span></span>

<span data-ttu-id="94e6b-108">Suivez ces meilleures pratiques lors de la personnalisation de la première expérience d’utilisation :</span><span class="sxs-lookup"><span data-stu-id="94e6b-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="94e6b-109">À faire</span><span class="sxs-lookup"><span data-stu-id="94e6b-109">Do</span></span>|<span data-ttu-id="94e6b-110">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="94e6b-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="94e6b-111">Proposer une simple et courte introduction aux actions principales disponibles dans le complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="94e6b-112">Ne pas inclure des informations et des détails qui ne sont pas pertinents pour la prise en main.</span><span class="sxs-lookup"><span data-stu-id="94e6b-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="94e6b-113">Donner aux utilisateurs la possibilité d’effectuer une action qui aura un impact positif sur leur utilisation du complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="94e6b-114">Ne pas espérer que les utilisateurs découvrent tous les éléments en même temps.</span><span class="sxs-lookup"><span data-stu-id="94e6b-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="94e6b-115">Concentrer les efforts sur le type ’action qui fournit le meilleur rendement.</span><span class="sxs-lookup"><span data-stu-id="94e6b-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="94e6b-116">Créer une expérience utilisateur attrayante que les utilisateurs vont vouloir compléter.</span><span class="sxs-lookup"><span data-stu-id="94e6b-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="94e6b-117">Ne pas forcer les utilisateurs à parcourir toute l’expérience de première utilisation.</span><span class="sxs-lookup"><span data-stu-id="94e6b-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="94e6b-118">Donner aux utilisateurs une option leur permettant d’ignorer l’expérience de première exécution.</span><span class="sxs-lookup"><span data-stu-id="94e6b-118">Give users an option to bypass the first-run experience.</span></span> |

<span data-ttu-id="94e6b-119">Déterminer s’il convient de montrer l’expérience de première utilisation une fois ou plusieurs fois (tout dépend de son importance pour votre scénario).</span><span class="sxs-lookup"><span data-stu-id="94e6b-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="94e6b-120">Par exemple, si votre complément est uniquement utilisé de temps en temps, les utilisateurs peuvent devenir moins familiarisés avec le complément. Ils pourraient alors bénéficier d’une autre interaction avec l’expérience de première exécution.</span><span class="sxs-lookup"><span data-stu-id="94e6b-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>

<span data-ttu-id="94e6b-121">Appliquer les modèles suivants le cas échéant pour créer ou optimisez l’expérience de première exécution de votre complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>

## <a name="carousel"></a><span data-ttu-id="94e6b-122">Carrousel</span><span class="sxs-lookup"><span data-stu-id="94e6b-122">Carousel</span></span>

<span data-ttu-id="94e6b-123">Le carrousel présente aux utilisateurs une série de fonctionnalités ou d’informations avant qu’ils ne commencent à utiliser le complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="94e6b-124">*Figure 1. Autoriser les utilisateurs à faire avancer ou ignorer les pages de début du flux carrousel*</span><span class="sxs-lookup"><span data-stu-id="94e6b-124">*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*</span></span>

![Illustration montrant l’étape 1 d’un carrousel lors de la première utilisation d’un volet Office’application de bureau.](../images/add-in-FRE-step-1.png)

<span data-ttu-id="94e6b-127">*Figure 2. Réduire le nombre d’écrans carrousels uniquement à ce qui est nécessaire pour communiquer efficacement votre message*</span><span class="sxs-lookup"><span data-stu-id="94e6b-127">*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*</span></span>

![Illustration montrant l’étape 2 d’un carrousel lors de la première expérience d’utilisation d’un Office d’application de bureau.](../images/add-in-FRE-step-2.png)

<span data-ttu-id="94e6b-130">*Figure 3. Fournir un appel clair à l’action pour quitter l’expérience de première run*</span><span class="sxs-lookup"><span data-stu-id="94e6b-130">*Figure 3. Provide a clear call to action to exit the first-run-experience*</span></span>

![Illustration montrant l’étape 3 d’un carrousel lors de la première expérience d’utilisation d’un Office d’application de bureau.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a><span data-ttu-id="94e6b-133">Mise en place de la valeur</span><span class="sxs-lookup"><span data-stu-id="94e6b-133">Value Placemat</span></span>

<span data-ttu-id="94e6b-134">La mise en place de la valeur communique la proposition de valeur de votre complément en faisant appel au positionnement de votre logo, à une proposition de valeur clairement déclarée, à une présentation ou un résumé des fonctionnalités et à une fonctionnalité claire d’appel à l’action.</span><span class="sxs-lookup"><span data-stu-id="94e6b-134">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>

<span data-ttu-id="94e6b-135">*Figure 4. Un lieu de valeur avec logo, proposition de valeur claire, résumé des fonctionnalités et appel à l’action*</span><span class="sxs-lookup"><span data-stu-id="94e6b-135">*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*</span></span>

![Illustration montrant une valeur de mise en place dans l’expérience de première utilisation d’un Office d’application de bureau.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a><span data-ttu-id="94e6b-138">Mise en place vidéo</span><span class="sxs-lookup"><span data-stu-id="94e6b-138">Video Placemat</span></span>

<span data-ttu-id="94e6b-139">La mise en place vidéo montre une vidéo aux utilisateurs avant qu’ils ne commencent à utiliser votre complément.</span><span class="sxs-lookup"><span data-stu-id="94e6b-139">The video placemat shows users a video before they start using your add-in.</span></span>

<span data-ttu-id="94e6b-140">*Figure 5. Première séquence de placemat vidéo - L’écran contient une image fixe de la vidéo avec un bouton lire et effacer le bouton d’appel à l’action*</span><span class="sxs-lookup"><span data-stu-id="94e6b-140">*Figure 5. First run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*</span></span>

![Illustration montrant une mise en place vidéo lors de la première expérience d’utilisation d’Office volet Des tâches de l’application de bureau.](../images/add-in-FRE-video.png)

<span data-ttu-id="94e6b-142">*Figure 6. Lecteur vidéo : les utilisateurs ont présenté une vidéo dans une fenêtre de boîte de dialogue*</span><span class="sxs-lookup"><span data-stu-id="94e6b-142">*Figure 6. Video player - Users presented with a video within a dialog window*</span></span>

![Illustration montrant une vidéo dans une fenêtre de boîte de dialogue avec une application Office de bureau et le volet Des tâches du add-in en arrière-plan.](../images/add-in-FRE-video-dialog.png)
