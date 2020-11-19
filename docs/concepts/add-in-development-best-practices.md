---
title: Meilleures pratiques en matière de développement de compléments Office
description: Appliquer les meilleures pratiques lors du développement pour créer des compléments Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 17393d921129efcfb74eed3dd168633c2f58291b
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132178"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="e151b-103">Meilleures pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="e151b-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="e151b-p101">Des compléments efficaces proposent des fonctionnalités uniques et attrayantes qui étendent les applications Office d’une manière visuellement attractive. Pour créer un complément intéressant, offrez une première expérience attractive à vos utilisateurs, concevez une interface utilisateur de premier choix et optimisez les performances de votre complément. Appliquez les meilleures pratiques décrites dans cet article pour créer des compléments permettant aux utilisateurs d’accomplir leurs tâches rapidement et efficacement.</span><span class="sxs-lookup"><span data-stu-id="e151b-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="e151b-107">Indication d’une valeur claire</span><span class="sxs-lookup"><span data-stu-id="e151b-107">Provide clear value</span></span>

- <span data-ttu-id="e151b-p102">Créez des compléments qui aident les utilisateurs à réaliser des tâches rapidement et efficacement. Concentrez-vous sur des scénarios adaptés aux applications Office. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="e151b-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
  - <span data-ttu-id="e151b-111">Réalisez des tâches de création essentielles plus rapidement et plus facilement, avec moins d’interruptions.</span><span class="sxs-lookup"><span data-stu-id="e151b-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
  - <span data-ttu-id="e151b-112">Développez de nouveaux scénarios dans Office.</span><span class="sxs-lookup"><span data-stu-id="e151b-112">Enable new scenarios within Office.</span></span>
  - <span data-ttu-id="e151b-113">Intégrez des services complémentaires dans les applications Office.</span><span class="sxs-lookup"><span data-stu-id="e151b-113">Embed complementary services within Office applications.</span></span>
  - <span data-ttu-id="e151b-114">Améliorez l’expérience Office pour accroître la productivité.</span><span class="sxs-lookup"><span data-stu-id="e151b-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="e151b-115">Assurez-vous que la valeur de votre complément apparaîtra clairement aux utilisateurs dès la première utilisation en créant une [première expérience enrichissante](#create-an-engaging-first-run-experience).</span><span class="sxs-lookup"><span data-stu-id="e151b-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="e151b-p103">Rédigez une [description claire pour AppSource](/office/dev/store/create-effective-office-store-listings). Soulignez les avantages de votre complément dans votre titre et votre description. Ne comptez pas sur votre marque pour communiquer sur les fonctionnalités de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e151b-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>

## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="e151b-119">Création d’une première expérience intéressante</span><span class="sxs-lookup"><span data-stu-id="e151b-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="e151b-p104">Attirez de nouveaux utilisateurs avec une première expérience très simple et intuitive. Les utilisateurs décident toujours d’utiliser ou d’abandonner un complément après l’avoir téléchargé à partir du Windows Store.</span><span class="sxs-lookup"><span data-stu-id="e151b-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="e151b-p105">Indiquez clairement les étapes que l’utilisateur doit suivre pour utiliser votre complément. Utilisez des vidéos, des schémas, des panneaux de pagination ou d’autres ressources pour attirer les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="e151b-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="e151b-124">N’hésitez pas à ajouter un texte pour insister sur l’utilité de votre complément sur l’écran de connexion des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="e151b-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="e151b-125">Proposez une interface utilisateur pédagogique pour guider les utilisateurs et la personnaliser.</span><span class="sxs-lookup"><span data-stu-id="e151b-125">Provide teaching UI to guide users and make your UI personal.</span></span>

  ![Capture d’écran illustrant une comparaison « do » vs.](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="e151b-129">Si votre complément de contenu est lié à des données dans le document de l’utilisateur, incluez des exemples de données ou un modèle pour montrer aux utilisateurs le format de données à utiliser.</span><span class="sxs-lookup"><span data-stu-id="e151b-129">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

  ![Capture d’écran illustrant une comparaison « do » vs.](../images/add-in-title.png)

- <span data-ttu-id="e151b-p108">Offrez des [essais gratuits](/office/dev/store/decide-on-a-pricing-model). Si votre complément nécessite un abonnement, proposez certaines fonctionnalités gratuitement.</span><span class="sxs-lookup"><span data-stu-id="e151b-p108">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="e151b-p109">Facilitez l’inscription. Préremplissez les informations (e-mail, nom d’affichage) et ignorez les vérifications d’adresses e-mail.</span><span class="sxs-lookup"><span data-stu-id="e151b-p109">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="e151b-p110">Évitez d’utiliser des fenêtres contextuelles. Si vous devez les utiliser, aidez les utilisateurs à les activer.</span><span class="sxs-lookup"><span data-stu-id="e151b-p110">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="e151b-139">Pour les modèles de conception à appliquer lors du développement de votre première expérience d’utilisation, reportez-vous à la section [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/first-run-experience-patterns.md).</span><span class="sxs-lookup"><span data-stu-id="e151b-139">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="e151b-140">Utilisation des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="e151b-140">Use add-in commands</span></span>

- <span data-ttu-id="e151b-p111">Fournissez des points d’entrée d’interface utilisateur pertinents pour votre complément à l’aide des commandes de complément. Pour plus d’informations, y compris les bonnes pratiques de conception, reportez-vous aux [commandes de complément](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="e151b-p111">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="e151b-143">Application des principes de conception de l’expérience utilisateur</span><span class="sxs-lookup"><span data-stu-id="e151b-143">Apply UX design principles</span></span>

- <span data-ttu-id="e151b-p112">Assurez-vous que l’aspect, la convivialité et la fonctionnalité de votre complément améliorent l’expérience Office. Utilisez [Office UI Fabric](https://developer.microsoft.com/fabric).</span><span class="sxs-lookup"><span data-stu-id="e151b-p112">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="e151b-p113">Privilégiez le contenu plutôt que l’apparence. Évitez les éléments d’interface utilisateur superflus qui n’ajoutent pas de valeur à l’expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e151b-p113">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="e151b-p114">Gardez le contrôle des utilisateurs. Assurez-vous que ces derniers comprennent les décisions importantes et peuvent facilement rétablir des actions effectuées par le complément.</span><span class="sxs-lookup"><span data-stu-id="e151b-p114">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="e151b-p115">Utilisez la personnalisation afin d’inspirer la confiance et d’orienter les utilisateurs. N’utilisez pas la personnalisation afin de submerger les utilisateurs ou de faire de la publicité.</span><span class="sxs-lookup"><span data-stu-id="e151b-p115">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="e151b-p116">Évitez d’utiliser le défilement. Optimisez votre complément pour une résolution de 1366 x 768.</span><span class="sxs-lookup"><span data-stu-id="e151b-p116">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="e151b-154">N’incluez pas d’image sans licence.</span><span class="sxs-lookup"><span data-stu-id="e151b-154">Do not include unlicensed images.</span></span>

- <span data-ttu-id="e151b-155">Utilisez un [langage clair et simple](../design/voice-guidelines.md) dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="e151b-155">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="e151b-156">Soulignez l’accessibilité : votre complément doit être facile à utiliser pour tous les utilisateurs et s’accommoder de technologies d’assistance telles que les lecteurs d’écran.</span><span class="sxs-lookup"><span data-stu-id="e151b-156">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="e151b-p117">Adaptez-le à toutes les plateformes et méthodes d’entrée, y compris la souris/le clavier et la [fonction tactile](#optimize-for-touch). Assurez-vous que votre interface utilisateur réagit à différents formats.</span><span class="sxs-lookup"><span data-stu-id="e151b-p117">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="e151b-159">Optimisation de la fonction tactile</span><span class="sxs-lookup"><span data-stu-id="e151b-159">Optimize for touch</span></span>

- <span data-ttu-id="e151b-160">Utilisez la propriété [Context. touchEnabled](/javascript/api/office/office.context#touchenabled) pour déterminer si l’application Office sur laquelle votre complément est exécuté est compatible avec la fonction tactile.</span><span class="sxs-lookup"><span data-stu-id="e151b-160">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="e151b-161">Cette propriété n’est pas prise en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="e151b-161">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="e151b-p118">Assurez-vous que toutes les commandes sont correctement dimensionnées pour l’interaction tactile. Par exemple, vérifiez que les boutons disposent de cibles tactiles adéquates et que les zones de texte sont assez grandes pour permettre la saisie.</span><span class="sxs-lookup"><span data-stu-id="e151b-p118">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="e151b-164">N’utilisez pas de méthodes d’entrée non tactiles comme l’utilisation du curseur ou du clic droit.</span><span class="sxs-lookup"><span data-stu-id="e151b-164">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="e151b-p119">Assurez-vous que votre complément fonctionne dans les modes portrait et paysage. Gardez à l’esprit qu’une partie de votre complément pourrait être masquée par le clavier virtuel sur les appareils tactiles.</span><span class="sxs-lookup"><span data-stu-id="e151b-p119">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="e151b-167">Testez votre complément sur un véritable appareil en utilisant le [chargement de version test](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="e151b-167">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="e151b-168">Si vous utilisez [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) pour vos éléments de conception, un grand nombre de ces éléments sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="e151b-168">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="e151b-169">Optimisation et contrôle des performances du complément</span><span class="sxs-lookup"><span data-stu-id="e151b-169">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="e151b-p120">Donnez l’impression que l’interface utilisateur réagit rapidement. Votre complément doit se charger en 500 ms au maximum.</span><span class="sxs-lookup"><span data-stu-id="e151b-p120">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="e151b-172">Veillez à ce que toutes les interactions utilisateur répondent en moins d’une seconde.</span><span class="sxs-lookup"><span data-stu-id="e151b-172">Ensure that all user interactions respond in under one second.</span></span>

- <span data-ttu-id="e151b-173">Fournissez des indicateurs de chargement pour les opérations à longue durée d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e151b-173">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="e151b-p121">Utilisez un CDN pour héberger les images, les ressources et les bibliothèques communes. Chargez autant d’éléments que possible à partir d’un seul emplacement.</span><span class="sxs-lookup"><span data-stu-id="e151b-p121">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="e151b-p122">Suivez les pratiques web standard pour optimiser votre page web. En production, utilisez uniquement les versions réduites des bibliothèques. Chargez uniquement les ressources dont vous avez besoin et optimisez leur chargement.</span><span class="sxs-lookup"><span data-stu-id="e151b-p122">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="e151b-p123">Si l’exécution des opérations dure longtemps, fournissez des commentaires aux utilisateurs. Prenez en compte les seuils indiqués dans le tableau suivant. Pour plus d’informations, reportez-vous à l’article sur les [limites des ressources et l’optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md).</span><span class="sxs-lookup"><span data-stu-id="e151b-p123">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="e151b-182">Classe d’interaction</span><span class="sxs-lookup"><span data-stu-id="e151b-182">Interaction class</span></span>|<span data-ttu-id="e151b-183">Target</span><span class="sxs-lookup"><span data-stu-id="e151b-183">Target</span></span>|<span data-ttu-id="e151b-184">Limite supérieure</span><span class="sxs-lookup"><span data-stu-id="e151b-184">Upper bound</span></span>|<span data-ttu-id="e151b-185">Perception humaine</span><span class="sxs-lookup"><span data-stu-id="e151b-185">Human perception</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="e151b-186">Instantanée</span><span class="sxs-lookup"><span data-stu-id="e151b-186">Instant</span></span>|<span data-ttu-id="e151b-187"><= 50 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-187"><=50 ms</span></span>|<span data-ttu-id="e151b-188">100 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-188">100 ms</span></span>|<span data-ttu-id="e151b-189">Aucun délai notable.</span><span class="sxs-lookup"><span data-stu-id="e151b-189">No noticeable delay.</span></span>|
  |<span data-ttu-id="e151b-190">Rapide</span><span class="sxs-lookup"><span data-stu-id="e151b-190">Fast</span></span>|<span data-ttu-id="e151b-191">50-100 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-191">50-100 ms</span></span>|<span data-ttu-id="e151b-192">200 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-192">200 ms</span></span>|<span data-ttu-id="e151b-p124">Délai notable minime. Aucun commentaire n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e151b-p124">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="e151b-195">Normale</span><span class="sxs-lookup"><span data-stu-id="e151b-195">Typical</span></span>|<span data-ttu-id="e151b-196">100-300 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-196">100-300 ms</span></span>|<span data-ttu-id="e151b-197">500 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-197">500 ms</span></span>|<span data-ttu-id="e151b-p125">L’opération va assez vite, sans pour autant pouvoir être qualifiée de rapide. Aucun commentaire n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e151b-p125">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="e151b-200">Réactive</span><span class="sxs-lookup"><span data-stu-id="e151b-200">Responsive</span></span>|<span data-ttu-id="e151b-201">300-500 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-201">300-500 ms</span></span>|<span data-ttu-id="e151b-202">1 seconde</span><span class="sxs-lookup"><span data-stu-id="e151b-202">1 second</span></span>|<span data-ttu-id="e151b-p126">L’opération n’est pas rapide, mais le système donne l’impression de répondre. Aucun commentaire n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e151b-p126">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="e151b-205">Continue</span><span class="sxs-lookup"><span data-stu-id="e151b-205">Continuous</span></span>|<span data-ttu-id="e151b-206">> 500 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-206">>500 ms</span></span>|<span data-ttu-id="e151b-207">5 secondes</span><span class="sxs-lookup"><span data-stu-id="e151b-207">5 seconds</span></span>|<span data-ttu-id="e151b-p127">Attente moyenne, le système n’a plus l’air de répondre. Un commentaire peut-être nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e151b-p127">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="e151b-210">Captive</span><span class="sxs-lookup"><span data-stu-id="e151b-210">Captive</span></span>|<span data-ttu-id="e151b-211">> 500 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-211">>500 ms</span></span>|<span data-ttu-id="e151b-212">10 secondes</span><span class="sxs-lookup"><span data-stu-id="e151b-212">10 seconds</span></span>|<span data-ttu-id="e151b-p128">Long, mais pas assez pour faire autre chose. Un commentaire peut-être nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e151b-p128">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="e151b-215">Étendue</span><span class="sxs-lookup"><span data-stu-id="e151b-215">Extended</span></span>|<span data-ttu-id="e151b-216">> 500 ms</span><span class="sxs-lookup"><span data-stu-id="e151b-216">>500 ms</span></span>|<span data-ttu-id="e151b-217">> 10 secondes</span><span class="sxs-lookup"><span data-stu-id="e151b-217">>10 seconds</span></span>|<span data-ttu-id="e151b-p129">Assez long pour faire quelque chose en attendant. Un commentaire peut être nécessaire.</span><span class="sxs-lookup"><span data-stu-id="e151b-p129">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="e151b-220">Longue durée d’exécution</span><span class="sxs-lookup"><span data-stu-id="e151b-220">Long running</span></span>|<span data-ttu-id="e151b-221">>5 secondes</span><span class="sxs-lookup"><span data-stu-id="e151b-221">>5 seconds</span></span>|<span data-ttu-id="e151b-222">>1 minute</span><span class="sxs-lookup"><span data-stu-id="e151b-222">>1 minute</span></span>|<span data-ttu-id="e151b-223">Les utilisateurs effectueront certainement une autre action.</span><span class="sxs-lookup"><span data-stu-id="e151b-223">Users will certainly do something else.</span></span>|

- <span data-ttu-id="e151b-224">Surveillez l’état de votre service et utilisez la télémétrie pour surveiller le succès d’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e151b-224">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="e151b-225">Réduisez les échanges de données entre le complément et le document Office.</span><span class="sxs-lookup"><span data-stu-id="e151b-225">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="e151b-226">Pour plus d’informations, reportez-vous à [la rubrique éviter d’utiliser la méthode Context. Sync dans les boucles](correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="e151b-226">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="e151b-227">Commercialisation de votre complément</span><span class="sxs-lookup"><span data-stu-id="e151b-227">Market your add-in</span></span>

- <span data-ttu-id="e151b-p131">Publiez votre complément dans [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) et [faites sa promotion](/office/dev/store/promote-your-office-store-solution) sur votre site web. Créez un [référencement AppSource efficace](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="e151b-p131">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="e151b-p132">Utilisez des titres et des descriptifs courts pour le complément. Ils ne doivent pas comporter plus de 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="e151b-p132">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="e151b-p133">Rédigez des descriptions brèves et attrayantes pour votre complément. Répondez à la question « Quel problème ce complément résout-il ? ».</span><span class="sxs-lookup"><span data-stu-id="e151b-p133">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="e151b-p134">Faites ressortir la proposition de valeur de votre complément dans le titre et la description. Ne comptez pas sur votre marque.</span><span class="sxs-lookup"><span data-stu-id="e151b-p134">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="e151b-236">Créez un site web pour aider les utilisateurs à trouver votre complément et à l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="e151b-236">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="e151b-237">Utiliser JavaScript qui prend en charge Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="e151b-237">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="e151b-238">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e151b-238">See also</span></span>

- [<span data-ttu-id="e151b-239">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="e151b-239">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="e151b-240">Découvrez le programme pour les développeurs Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="e151b-240">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
