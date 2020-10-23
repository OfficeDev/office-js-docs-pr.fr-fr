---
title: Meilleures pratiques en matière de développement de compléments Office
description: Appliquer les meilleures pratiques lors du développement pour créer des compléments Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 8ce0482e108e7b8774442a2b0669a0e76bb401f9
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740860"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="88a93-103">Meilleures pratiques en matière de développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="88a93-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="88a93-p101">Des compléments efficaces proposent des fonctionnalités uniques et attrayantes qui étendent les applications Office d’une manière visuellement attractive. Pour créer un complément intéressant, offrez une première expérience attractive à vos utilisateurs, concevez une interface utilisateur de premier choix et optimisez les performances de votre complément. Appliquez les meilleures pratiques décrites dans cet article pour créer des compléments permettant aux utilisateurs d’accomplir leurs tâches rapidement et efficacement.</span><span class="sxs-lookup"><span data-stu-id="88a93-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="88a93-107">Indication d’une valeur claire</span><span class="sxs-lookup"><span data-stu-id="88a93-107">Provide clear value</span></span>

- <span data-ttu-id="88a93-p102">Créez des compléments qui aident les utilisateurs à réaliser des tâches rapidement et efficacement. Concentrez-vous sur des scénarios adaptés aux applications Office. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="88a93-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
 - <span data-ttu-id="88a93-111">Réalisez des tâches de création essentielles plus rapidement et plus facilement, avec moins d’interruptions.</span><span class="sxs-lookup"><span data-stu-id="88a93-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
 - <span data-ttu-id="88a93-112">Développez de nouveaux scénarios dans Office.</span><span class="sxs-lookup"><span data-stu-id="88a93-112">Enable new scenarios within Office.</span></span>
 - <span data-ttu-id="88a93-113">Intégrez des services complémentaires dans les applications Office.</span><span class="sxs-lookup"><span data-stu-id="88a93-113">Embed complementary services within Office applications.</span></span>
 - <span data-ttu-id="88a93-114">Améliorez l’expérience Office pour accroître la productivité.</span><span class="sxs-lookup"><span data-stu-id="88a93-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="88a93-115">Assurez-vous que la valeur de votre complément apparaîtra clairement aux utilisateurs dès la première utilisation en créant une [première expérience enrichissante](#create-an-engaging-first-run-experience).</span><span class="sxs-lookup"><span data-stu-id="88a93-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="88a93-p103">Rédigez une [description claire pour AppSource](/office/dev/store/create-effective-office-store-listings). Soulignez les avantages de votre complément dans votre titre et votre description. Ne comptez pas sur votre marque pour communiquer sur les fonctionnalités de votre complément.</span><span class="sxs-lookup"><span data-stu-id="88a93-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>


## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="88a93-119">Création d’une première expérience intéressante</span><span class="sxs-lookup"><span data-stu-id="88a93-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="88a93-p104">Attirez de nouveaux utilisateurs avec une première expérience très simple et intuitive. Les utilisateurs décident toujours d’utiliser ou d’abandonner un complément après l’avoir téléchargé à partir du Windows Store.</span><span class="sxs-lookup"><span data-stu-id="88a93-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="88a93-p105">Indiquez clairement les étapes que l’utilisateur doit suivre pour utiliser votre complément. Utilisez des vidéos, des schémas, des panneaux de pagination ou d’autres ressources pour attirer les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="88a93-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="88a93-124">N’hésitez pas à ajouter un texte pour insister sur l’utilité de votre complément sur l’écran de connexion des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="88a93-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="88a93-125">Proposez une interface utilisateur pédagogique pour guider les utilisateurs et la personnaliser.</span><span class="sxs-lookup"><span data-stu-id="88a93-125">Provide teaching UI to guide users and make your UI personal.</span></span>

   ![Capture d’écran illustrant un complément de volet Office avec des étapes de mise en route en regard d’un complément sans étapes de mise en route](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="88a93-127">Si votre complément de contenu est lié à des données dans le document de l’utilisateur, incluez des exemples de données ou un modèle pour montrer aux utilisateurs le format de données à utiliser.</span><span class="sxs-lookup"><span data-stu-id="88a93-127">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

   ![Capture d’écran illustrant un complément de contenu avec des données en regard d’un complément de contenu sans données](../images/add-in-title.png)

- <span data-ttu-id="88a93-p106">Offrez des [essais gratuits](/office/dev/store/decide-on-a-pricing-model). Si votre complément nécessite un abonnement, proposez certaines fonctionnalités gratuitement.</span><span class="sxs-lookup"><span data-stu-id="88a93-p106">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="88a93-p107">Facilitez l’inscription. Préremplissez les informations (e-mail, nom d’affichage) et ignorez les vérifications d’adresses e-mail.</span><span class="sxs-lookup"><span data-stu-id="88a93-p107">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="88a93-p108">Évitez d’utiliser des fenêtres contextuelles. Si vous devez les utiliser, aidez les utilisateurs à les activer.</span><span class="sxs-lookup"><span data-stu-id="88a93-p108">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="88a93-135">Pour les modèles de conception à appliquer lors du développement de votre première expérience d’utilisation, reportez-vous à la section [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/first-run-experience-patterns.md).</span><span class="sxs-lookup"><span data-stu-id="88a93-135">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="88a93-136">Utilisation des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="88a93-136">Use add-in commands</span></span>

- <span data-ttu-id="88a93-p109">Fournissez des points d’entrée d’interface utilisateur pertinents pour votre complément à l’aide des commandes de complément. Pour plus d’informations, y compris les bonnes pratiques de conception, reportez-vous aux [commandes de complément](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="88a93-p109">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="88a93-139">Application des principes de conception de l’expérience utilisateur</span><span class="sxs-lookup"><span data-stu-id="88a93-139">Apply UX design principles</span></span>

- <span data-ttu-id="88a93-p110">Assurez-vous que l’aspect, la convivialité et la fonctionnalité de votre complément améliorent l’expérience Office. Utilisez [Office UI Fabric](https://developer.microsoft.com/fabric).</span><span class="sxs-lookup"><span data-stu-id="88a93-p110">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="88a93-p111">Privilégiez le contenu plutôt que l’apparence. Évitez les éléments d’interface utilisateur superflus qui n’ajoutent pas de valeur à l’expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="88a93-p111">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="88a93-p112">Gardez le contrôle des utilisateurs. Assurez-vous que ces derniers comprennent les décisions importantes et peuvent facilement rétablir des actions effectuées par le complément.</span><span class="sxs-lookup"><span data-stu-id="88a93-p112">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="88a93-p113">Utilisez la personnalisation afin d’inspirer la confiance et d’orienter les utilisateurs. N’utilisez pas la personnalisation afin de submerger les utilisateurs ou de faire de la publicité.</span><span class="sxs-lookup"><span data-stu-id="88a93-p113">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="88a93-p114">Évitez d’utiliser le défilement. Optimisez votre complément pour une résolution de 1366 x 768.</span><span class="sxs-lookup"><span data-stu-id="88a93-p114">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="88a93-150">N’incluez pas d’image sans licence.</span><span class="sxs-lookup"><span data-stu-id="88a93-150">Do not include unlicensed images.</span></span>

- <span data-ttu-id="88a93-151">Utilisez un [langage clair et simple](../design/voice-guidelines.md) dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="88a93-151">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="88a93-152">Soulignez l’accessibilité : votre complément doit être facile à utiliser pour tous les utilisateurs et s’accommoder de technologies d’assistance telles que les lecteurs d’écran.</span><span class="sxs-lookup"><span data-stu-id="88a93-152">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="88a93-p115">Adaptez-le à toutes les plateformes et méthodes d’entrée, y compris la souris/le clavier et la [fonction tactile](#optimize-for-touch). Assurez-vous que votre interface utilisateur réagit à différents formats.</span><span class="sxs-lookup"><span data-stu-id="88a93-p115">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="88a93-155">Optimisation de la fonction tactile</span><span class="sxs-lookup"><span data-stu-id="88a93-155">Optimize for touch</span></span>

- <span data-ttu-id="88a93-156">Utilisez la propriété [Context. touchEnabled](/javascript/api/office/office.context#touchenabled) pour déterminer si l’application Office sur laquelle votre complément est exécuté est compatible avec la fonction tactile.</span><span class="sxs-lookup"><span data-stu-id="88a93-156">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="88a93-157">Cette propriété n’est pas prise en charge dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="88a93-157">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="88a93-p116">Assurez-vous que toutes les commandes sont correctement dimensionnées pour l’interaction tactile. Par exemple, vérifiez que les boutons disposent de cibles tactiles adéquates et que les zones de texte sont assez grandes pour permettre la saisie.</span><span class="sxs-lookup"><span data-stu-id="88a93-p116">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="88a93-160">N’utilisez pas de méthodes d’entrée non tactiles comme l’utilisation du curseur ou du clic droit.</span><span class="sxs-lookup"><span data-stu-id="88a93-160">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="88a93-p117">Assurez-vous que votre complément fonctionne dans les modes portrait et paysage. Gardez à l’esprit qu’une partie de votre complément pourrait être masquée par le clavier virtuel sur les appareils tactiles.</span><span class="sxs-lookup"><span data-stu-id="88a93-p117">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="88a93-163">Testez votre complément sur un véritable appareil en utilisant le [chargement de version test](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="88a93-163">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="88a93-164">Si vous utilisez [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) pour vos éléments de conception, un grand nombre de ces éléments sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="88a93-164">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="88a93-165">Optimisation et contrôle des performances du complément</span><span class="sxs-lookup"><span data-stu-id="88a93-165">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="88a93-p118">Donnez l’impression que l’interface utilisateur réagit rapidement. Votre complément doit se charger en 500 ms au maximum.</span><span class="sxs-lookup"><span data-stu-id="88a93-p118">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="88a93-168">Veillez à ce que toutes les interactions utilisateur répondent en moins d’une seconde.</span><span class="sxs-lookup"><span data-stu-id="88a93-168">Ensure that all user interactions respond in under one second.</span></span>

-  <span data-ttu-id="88a93-169">Fournissez des indicateurs de chargement pour les opérations à longue durée d’exécution.</span><span class="sxs-lookup"><span data-stu-id="88a93-169">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="88a93-p119">Utilisez un CDN pour héberger les images, les ressources et les bibliothèques communes. Chargez autant d’éléments que possible à partir d’un seul emplacement.</span><span class="sxs-lookup"><span data-stu-id="88a93-p119">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="88a93-p120">Suivez les pratiques web standard pour optimiser votre page web. En production, utilisez uniquement les versions réduites des bibliothèques. Chargez uniquement les ressources dont vous avez besoin et optimisez leur chargement.</span><span class="sxs-lookup"><span data-stu-id="88a93-p120">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="88a93-p121">Si l’exécution des opérations dure longtemps, fournissez des commentaires aux utilisateurs. Prenez en compte les seuils indiqués dans le tableau suivant. Pour plus d’informations, reportez-vous à l’article sur les [limites des ressources et l’optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md).</span><span class="sxs-lookup"><span data-stu-id="88a93-p121">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="88a93-178">**Classe d’interaction**</span><span class="sxs-lookup"><span data-stu-id="88a93-178">**Interaction class**</span></span>|<span data-ttu-id="88a93-179">**Cible**</span><span class="sxs-lookup"><span data-stu-id="88a93-179">**Target**</span></span>|<span data-ttu-id="88a93-180">**Limite supérieure**</span><span class="sxs-lookup"><span data-stu-id="88a93-180">**Upper bound**</span></span>|<span data-ttu-id="88a93-181">**Perception humaine**</span><span class="sxs-lookup"><span data-stu-id="88a93-181">**Human perception**</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="88a93-182">Instantanée</span><span class="sxs-lookup"><span data-stu-id="88a93-182">Instant</span></span>|<span data-ttu-id="88a93-183"><= 50 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-183"><=50 ms</span></span>|<span data-ttu-id="88a93-184">100 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-184">100 ms</span></span>|<span data-ttu-id="88a93-185">Aucun délai notable.</span><span class="sxs-lookup"><span data-stu-id="88a93-185">No noticeable delay.</span></span>|
  |<span data-ttu-id="88a93-186">Rapide</span><span class="sxs-lookup"><span data-stu-id="88a93-186">Fast</span></span>|<span data-ttu-id="88a93-187">50-100 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-187">50-100 ms</span></span>|<span data-ttu-id="88a93-188">200 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-188">200 ms</span></span>|<span data-ttu-id="88a93-p122">Délai notable minime. Aucun commentaire n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="88a93-p122">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="88a93-191">Normale</span><span class="sxs-lookup"><span data-stu-id="88a93-191">Typical</span></span>|<span data-ttu-id="88a93-192">100-300 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-192">100-300 ms</span></span>|<span data-ttu-id="88a93-193">500 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-193">500 ms</span></span>|<span data-ttu-id="88a93-p123">L’opération va assez vite, sans pour autant pouvoir être qualifiée de rapide. Aucun commentaire n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="88a93-p123">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="88a93-196">Réactive</span><span class="sxs-lookup"><span data-stu-id="88a93-196">Responsive</span></span>|<span data-ttu-id="88a93-197">300-500 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-197">300-500 ms</span></span>|<span data-ttu-id="88a93-198">1 seconde</span><span class="sxs-lookup"><span data-stu-id="88a93-198">1 second</span></span>|<span data-ttu-id="88a93-p124">L’opération n’est pas rapide, mais le système donne l’impression de répondre. Aucun commentaire n’est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="88a93-p124">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="88a93-201">Continue</span><span class="sxs-lookup"><span data-stu-id="88a93-201">Continuous</span></span>|<span data-ttu-id="88a93-202">> 500 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-202">>500 ms</span></span>|<span data-ttu-id="88a93-203">5 secondes</span><span class="sxs-lookup"><span data-stu-id="88a93-203">5 seconds</span></span>|<span data-ttu-id="88a93-p125">Attente moyenne, le système n’a plus l’air de répondre. Un commentaire peut-être nécessaire.</span><span class="sxs-lookup"><span data-stu-id="88a93-p125">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="88a93-206">Captive</span><span class="sxs-lookup"><span data-stu-id="88a93-206">Captive</span></span>|<span data-ttu-id="88a93-207">> 500 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-207">>500 ms</span></span>|<span data-ttu-id="88a93-208">10 secondes</span><span class="sxs-lookup"><span data-stu-id="88a93-208">10 seconds</span></span>|<span data-ttu-id="88a93-p126">Long, mais pas assez pour faire autre chose. Un commentaire peut-être nécessaire.</span><span class="sxs-lookup"><span data-stu-id="88a93-p126">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="88a93-211">Étendue</span><span class="sxs-lookup"><span data-stu-id="88a93-211">Extended</span></span>|<span data-ttu-id="88a93-212">> 500 ms</span><span class="sxs-lookup"><span data-stu-id="88a93-212">>500 ms</span></span>|<span data-ttu-id="88a93-213">> 10 secondes</span><span class="sxs-lookup"><span data-stu-id="88a93-213">>10 seconds</span></span>|<span data-ttu-id="88a93-p127">Assez long pour faire quelque chose en attendant. Un commentaire peut être nécessaire.</span><span class="sxs-lookup"><span data-stu-id="88a93-p127">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="88a93-216">Longue durée d’exécution</span><span class="sxs-lookup"><span data-stu-id="88a93-216">Long running</span></span>|<span data-ttu-id="88a93-217">>5 secondes</span><span class="sxs-lookup"><span data-stu-id="88a93-217">>5 seconds</span></span>|<span data-ttu-id="88a93-218">>1 minute</span><span class="sxs-lookup"><span data-stu-id="88a93-218">>1 minute</span></span>|<span data-ttu-id="88a93-219">Les utilisateurs effectueront certainement une autre action.</span><span class="sxs-lookup"><span data-stu-id="88a93-219">Users will certainly do something else.</span></span>|

- <span data-ttu-id="88a93-220">Surveillez l’état de votre service et utilisez la télémétrie pour surveiller le succès d’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="88a93-220">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="88a93-221">Réduisez les échanges de données entre le complément et le document Office.</span><span class="sxs-lookup"><span data-stu-id="88a93-221">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="88a93-222">Pour plus d’informations, reportez-vous à [la rubrique éviter d’utiliser la méthode Context. Sync dans les boucles](correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="88a93-222">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="88a93-223">Commercialisation de votre complément</span><span class="sxs-lookup"><span data-stu-id="88a93-223">Market your add-in</span></span>

- <span data-ttu-id="88a93-p129">Publiez votre complément dans [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) et [faites sa promotion](/office/dev/store/promote-your-office-store-solution) sur votre site web. Créez un [référencement AppSource efficace](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="88a93-p129">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="88a93-p130">Utilisez des titres et des descriptifs courts pour le complément. Ils ne doivent pas comporter plus de 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="88a93-p130">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="88a93-p131">Rédigez des descriptions brèves et attrayantes pour votre complément. Répondez à la question « Quel problème ce complément résout-il ? ».</span><span class="sxs-lookup"><span data-stu-id="88a93-p131">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="88a93-p132">Faites ressortir la proposition de valeur de votre complément dans le titre et la description. Ne comptez pas sur votre marque.</span><span class="sxs-lookup"><span data-stu-id="88a93-p132">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="88a93-232">Créez un site web pour aider les utilisateurs à trouver votre complément et à l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="88a93-232">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="88a93-233">Utiliser JavaScript qui prend en charge Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="88a93-233">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="88a93-234">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="88a93-234">See also</span></span>

- [<span data-ttu-id="88a93-235">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="88a93-235">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="88a93-236">En savoir plus sur le programme de développement Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="88a93-236">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
