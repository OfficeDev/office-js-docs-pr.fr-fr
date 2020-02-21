---
title: Conception des compléments Outlook
description: Les instructions suivantes vous aideront à concevoir et à créer un complément attrayant, qui apportera le meilleur de votre application directement dans Outlook sur Windows, le web, iOS, Mac et Android.
ms.date: 06/24/2019
localization_priority: Priority
ms.openlocfilehash: efedeb32643bff12e167931ac4da80fdcc2c277f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166099"
---
# <a name="outlook-add-in-design-guidelines"></a><span data-ttu-id="e304c-103">Instructions de création d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="e304c-103">Outlook add-in design guidelines</span></span>

<span data-ttu-id="e304c-p101">Les compléments sont un excellent moyen pour les partenaires d’étendre les fonctionnalités d’Outlook au-delà de notre ensemble de fonctionnalités de base. Les compléments permettent aux utilisateurs d’accéder à des expériences, des tâches et du contenu de tiers sans avoir à quitter leur boîte de réception. Une fois installés, les compléments Outlook sont disponibles sur toutes les plateformes et tous les appareils.</span><span class="sxs-lookup"><span data-stu-id="e304c-p101">Add-ins are a great way for partners to extend the functionality of Outlook beyond our core feature set. Add-ins enable users to access third-party experiences, tasks, and content without needing to leave their inbox. Once installed, Outlook add-ins are available on every platform and device.</span></span>  

<span data-ttu-id="e304c-107">Les instructions de haut niveau suivantes vous aideront à concevoir et à créer un complément attrayant, qui apportera le meilleur de votre application directement dans Outlook&mdash; sur Windows, le web, iOS, Mac et Android.</span><span class="sxs-lookup"><span data-stu-id="e304c-107">The following high-level guidelines will help you design and build a compelling add-in, which brings the best of your app right into Outlook&mdash;on Windows, Web, iOS, Mac, and Android.</span></span>

## <a name="principles"></a><span data-ttu-id="e304c-108">Principes</span><span class="sxs-lookup"><span data-stu-id="e304c-108">Principles</span></span>

1. <span data-ttu-id="e304c-109">**Concentrez-vous sur quelques tâches clés et exécutez-les correctement**</span><span class="sxs-lookup"><span data-stu-id="e304c-109">**Focus on a few key tasks; do them well**</span></span>

   <span data-ttu-id="e304c-p102">Les compléments les mieux conçus sont simples à utiliser, visent un objectif précis et sont réellement utiles pour les utilisateurs. Votre complément s’exécutera dans Outlook, ce principe est donc d’autant plus important. Outlook est une application de productivité&mdash;c’est l’endroit où les utilisateurs se rendent pour s’acquitter de leurs tâches.</span><span class="sxs-lookup"><span data-stu-id="e304c-p102">The best designed add-ins are simple to use, focused, and provide real value to users. Because your add-in will run inside of Outlook, there is additional emphasis placed on this principle. Outlook is a productivity app&mdash;it's where people go to get things done.</span></span>

   <span data-ttu-id="e304c-p103">Vous allez apporter une extension à notre expérience et vous devez être certain que les scénarios que vous activez s’intègre naturellement au sein d’Outlook. Réfléchissez bien aux situations dans lesquelles la présence des compléments sera le plus utile pour les utilisateurs dans les expériences de messagerie et de calendrier.</span><span class="sxs-lookup"><span data-stu-id="e304c-p103">You will be an extension of our experience and it is important to make sure the scenarios you enable feel like a natural fit inside of Outlook. Think carefully about which of your common use cases will benefit the most from having hooks to them from within our email and calendaring experiences.</span></span>

   <span data-ttu-id="e304c-p104">Un complément ne doit pas tenter d’exécuter tout ce que votre application fait déjà. Concentrez-vous sur les actions appropriées les plus fréquemment utilisées, dans le contexte de contenu Outlook. Pensez à votre appel à l’action et indiquez clairement à l’utilisateur ce qu’il doit faire lorsque votre volet de tâches s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="e304c-p104">An add-in should not attempt to do everything your app does. The focus should be on the most frequently used, and appropriate, actions in the context of Outlook content. Think about your call to action and make it clear what the user should do when your task pane opens.</span></span>

2. <span data-ttu-id="e304c-118">**Faites en sorte que tout semble aussi naturel que possible**</span><span class="sxs-lookup"><span data-stu-id="e304c-118">**Make it feel as native as possible**</span></span>

   <span data-ttu-id="e304c-p105">Votre complément doit être conçu à l’aide de schémas natifs de la plateforme sur laquelle Outlook s’exécute. Pour ce faire, veillez à respecter et implémenter les instructions d’interaction et visuelles définies par chaque plateforme. Outlook possède ses propres instructions et celles-ci doivent également être prises en compte. Un complément bien conçu sera une combinaison appropriée de votre expérience, de la plateforme et d’Outlook.</span><span class="sxs-lookup"><span data-stu-id="e304c-p105">Your add-in should be designed using patterns native to the platform that Outlook is running on. To achieve this, be sure to respect and implement the interaction and visual guidelines set forth by each platform. Outlook has its own guidelines and those are also important to consider. A well-designed add-in will be an appropriate blend of your experience, the platform, and Outlook.</span></span>

   <span data-ttu-id="e304c-p106">Cela ne signifie pas que votre complément devra être différent visuellement lorsqu’il est exécuté sur Outlook sur iOS et Outlook sur Android. Nous vous recommandons de vous référer à [Framework7](https://framework7.io/) comme une option pour vous aider dans les styles.</span><span class="sxs-lookup"><span data-stu-id="e304c-p106">This does mean that your add-in will have to visually be different when it runs in Outlook on iOS versus Android. We recommend taking a look at [Framework7](https://framework7.io/) as one option to help you with styling.</span></span>

3. <span data-ttu-id="e304c-125">**Faites en sorte que votre complément soit agréable à utiliser jusque dans les moindres détails**</span><span class="sxs-lookup"><span data-stu-id="e304c-125">**Make it enjoyable to use and get the details right**</span></span>

   <span data-ttu-id="e304c-p107">Les gens apprécient les produits qui sont attrayants sur le plan fonctionnel et visuel. Vous pouvez contribuer à garantir le succès de votre complément en créant une expérience où vous avez tenu soigneusement compte de chaque interaction et détail visuel. Les étapes nécessaires à l’exécution d’une tâche doivent être claires et pertinentes. Dans l’idéal, aucune action ne doit exiger plus d’un clic ou deux.</span><span class="sxs-lookup"><span data-stu-id="e304c-p107">People enjoy using products that are both functionally and visually appealing. You can help ensure the success of your add-in by crafting an experience where you've carefully considered every interaction and visual detail. The necessary steps to complete a task must be clear and relevant. Ideally, no action should be further than a click or two away.</span></span> 
   
   <span data-ttu-id="e304c-130">Un utilisateur ne doit pas sortir du contexte pertinent pour effectuer une action.</span><span class="sxs-lookup"><span data-stu-id="e304c-130">Try not to take a user out of context to complete an action.</span></span> <span data-ttu-id="e304c-131">Un utilisateur doit pouvoir accéder à votre complément et en sortir facilement pour revenir à ce qu’il faisait avant.</span><span class="sxs-lookup"><span data-stu-id="e304c-131">A user should easily be able to get in and out of your add-in and back to whatever she was doing before.</span></span> <span data-ttu-id="e304c-132">Un complément n’est pas destiné à être un emplacement où l’utilisateur passe beaucoup de temps&mdash;il doit s’agir d’une amélioration de nos fonctionnalités principales.</span><span class="sxs-lookup"><span data-stu-id="e304c-132">An add-in is not meant to be a destination to spend a lot of time in&mdash;it is an enhancement to our core functionality.</span></span> <span data-ttu-id="e304c-133">Si votre complément est développé correctement, il nous aidera à augmenter la productivité des utilisateurs, ce qui constitue un de nos objectifs.</span><span class="sxs-lookup"><span data-stu-id="e304c-133">If done properly, your add-in will help us deliver on the goal of making people more productive.</span></span>

4. <span data-ttu-id="e304c-134">**Personnalisez votre complément à l’image de votre marque de manière judicieuse**</span><span class="sxs-lookup"><span data-stu-id="e304c-134">**Brand wisely**</span></span>

   <span data-ttu-id="e304c-135">Nous apprécions les personnalisations et nous savons qu’il est important pour vous de procurer votre expérience unique aux utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="e304c-135">We value great branding, and we know it is important to provide users with your unique experience.</span></span> <span data-ttu-id="e304c-136">Cependant, nous pensons que la meilleure façon de garantir la réussite de votre complément est de créer une expérience intuitive qui incorpore subtilement les éléments de votre marque au lieu d’afficher des éléments de marque permanents ou obstruants qui empêchent les utilisateurs de naviguer dans votre système de manière fluide.</span><span class="sxs-lookup"><span data-stu-id="e304c-136">But we feel the best way to ensure your add-in's success is to build an intuitive experience that subtly incorporates elements of your brand versus displaying persistent or obtrusive brand elements that only distract a user from moving through your system in an unencumbered manner.</span></span> 
    
   <span data-ttu-id="e304c-137">Vous pouvez par exemple intégrer votre marque en utilisant les couleurs, les icônes et le ton qui la définissent&mdash;tout en respectant les modèles privilégiés de la plateforme et les critères d’accessibilité.</span><span class="sxs-lookup"><span data-stu-id="e304c-137">A good way to incorporate your brand in a meaningful way is through the use of your brand colors, icons, and voice&mdash;assuming these don't conflict with the preferred platform patterns or accessibility requirements.</span></span> <span data-ttu-id="e304c-138">Efforcez-vous de toujours privilégier le contenu et la capacité à effectuer des tâches plutôt que de chercher à attirer l’attention sur votre marque.</span><span class="sxs-lookup"><span data-stu-id="e304c-138">Strive to keep the focus on content and task completion, not brand attention.</span></span> 
    
   > [!NOTE]
   >  <span data-ttu-id="e304c-139">Les publicités ne doivent pas être affichées dans des compléments sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e304c-139">Ads should not be shown within add-ins on iOS or Android.</span></span>

## <a name="design-patterns"></a><span data-ttu-id="e304c-140">Modèles de conception</span><span class="sxs-lookup"><span data-stu-id="e304c-140">Design patterns</span></span>

> [!NOTE]
> <span data-ttu-id="e304c-141">Tandis que les principes ci-dessus s’appliquent à l’ensemble des points de terminaison/plateformes, les modèles et les exemples suivants sont spécifiques des compléments mobiles sur la plateforme iOS.</span><span class="sxs-lookup"><span data-stu-id="e304c-141">While the above principles apply to all endpoints/platforms, the following patterns and examples are specific to mobile add-ins on the iOS platform.</span></span>

<span data-ttu-id="e304c-p111">Pour vous aider à créer un complément bien conçu, nous proposons des [modèles](../design/ux-design-pattern-templates.md) pour les versions mobiles avec iOS fonctionnant dans l’environnement Outlook Mobile. Si vous utilisez ces modèles spécifiques, votre complément semblera natif de la plateforme iOS et d’Outlook Mobile. Ces modèles sont également décrits en détail ci-dessous. Bien que cette bibliothèque ne soit pas exhaustive, il s’agit du début de son développement et nous continuerons à l’enrichir à mesure que nous découvrirons des paradigmes que nos partenaires souhaitent inclure dans leurs compléments.</span><span class="sxs-lookup"><span data-stu-id="e304c-p111">To help you create a well-designed add-in, we have [templates](../design/ux-design-pattern-templates.md) that contain iOS mobile patterns that work within the Outlook Mobile environment. Leveraging these specific patterns will help ensure your add-in feels native to both the iOS platform and Outlook Mobile. These patterns are also detailed below. While not exhaustive, this is the start of a library that we will continue to build upon as we uncover additional paradigms partners wish to include in their add-ins.</span></span>  

### <a name="overview"></a><span data-ttu-id="e304c-146">Vue d’ensemble</span><span class="sxs-lookup"><span data-stu-id="e304c-146">Overview</span></span>

<span data-ttu-id="e304c-147">Un complément type est constitué des éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="e304c-147">A typical add-in is made up of the following components.</span></span>

![Diagramme de modèles d’expérience utilisateur de base pour un volet de tâches sur iOS](../images/outlook-mobile-design-overview.png)

![Diagramme de modèles d’expérience utilisateur de base pour un volet de tâches sur Android](../images/outlook-mobile-design-overview-android.jpg)

### <a name="loading"></a><span data-ttu-id="e304c-150">Chargement</span><span class="sxs-lookup"><span data-stu-id="e304c-150">Loading</span></span>

<span data-ttu-id="e304c-p112">Lorsqu’un utilisateur sélectionne votre complément, l’expérience utilisateur doit s’afficher rapidement. Si le chargement est long, utilisez une barre de progression ou un indicateur d’activité. Une barre de progression doit être utilisée lorsque le délai peut être déterminé et un indicateur d’activité doit être utilisé lorsque le délai ne peut pas être déterminé.</span><span class="sxs-lookup"><span data-stu-id="e304c-p112">When a user taps on your add-in, the UX should display as quickly as possible. If there is any delay, use a progress bar or activity indicator. A progress bar should be used when the amount of time is determinable and an activity indicator should be used when the amount of time is indeterminable.</span></span>

<span data-ttu-id="e304c-154">**Exemple de chargement de pages sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-154">**An example of loading pages on iOS**</span></span>

![Exemples illustrant une barre de progression et un indicateur d’activité sur iOS](../images/outlook-mobile-design-loading.png)

<span data-ttu-id="e304c-156">**Exemple de chargement de pages sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-156">**An example of loading pages on Android**</span></span>

![Exemples illustrant une barre de progression et un indicateur d’activité sur Android](../images/outlook-mobile-design-loading-android.jpg)


### <a name="sign-insign-up"></a><span data-ttu-id="e304c-158">Connexion/Inscription</span><span class="sxs-lookup"><span data-stu-id="e304c-158">Sign in/Sign up</span></span>

<span data-ttu-id="e304c-159">Votre procédure de connexion (et d’inscription) doit être directe et simple.</span><span class="sxs-lookup"><span data-stu-id="e304c-159">Make your sign in (and sign up) flow straightforward and simple to use.</span></span>

<span data-ttu-id="e304c-160">**Exemple de page de connexion et d’inscription sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-160">**An example sign in and sign up page on iOS**</span></span>

![Exemples de pages de connexion et d’inscription sur iOS](../images/outlook-mobile-design-signin.png)

<span data-ttu-id="e304c-162">**Exemple de page de connexion sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-162">**An example sign in page on Android**</span></span>

![Exemples de page de connexion sur Android](../images/outlook-mobile-design-signin-android.png)

### <a name="brand-bar"></a><span data-ttu-id="e304c-164">Barre de marque</span><span class="sxs-lookup"><span data-stu-id="e304c-164">Brand bar</span></span>

<span data-ttu-id="e304c-p113">Le premier écran de votre complément doit inclure un élément de votre marque. Conçue pour que votre marque soit reconnue, la barre de marque vous aide également à définir le contexte pour l’utilisateur. Étant donné que la barre de navigation contient le nom de votre société/marque, il est inutile de reproduire la barre de marque sur les pages suivantes.</span><span class="sxs-lookup"><span data-stu-id="e304c-p113">The first screen of your add-in should include your branding element. Designed for recognition, the brand bar also helps set context for the user. Because the navigation bar contains the name of your company/brand, it's unnecessary to repeat the brand bar on subsequent pages.</span></span>

<span data-ttu-id="e304c-168">**Exemple de personnalisation sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-168">**An example of branding on iOS**</span></span>

![Exemples de barres de marque sur iOS](../images/outlook-mobile-design-branding.png)

<span data-ttu-id="e304c-170">**Exemple de personnalisation sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-170">**An example of branding on Android**</span></span>

![Exemples de barres de marque sur Android](../images/outlook-mobile-design-branding-android.png)

### <a name="margins"></a><span data-ttu-id="e304c-172">Marges</span><span class="sxs-lookup"><span data-stu-id="e304c-172">Margins</span></span>

<span data-ttu-id="e304c-173">Les marges sur mobile doivent être définies sur 15 px (8 % de l’écran) pour chaque côté afin de s’aligner sur Outlook iOS et sur 16 px pour chaque côté afin de s’aligner sur Outlook Android.</span><span class="sxs-lookup"><span data-stu-id="e304c-173">Mobile margins should be set to 15px (8% of screen) for each side, to align with Outlook iOS and 16px for each side to align with Outlook Android.</span></span>

![Exemples de marges sur iOS](../images/outlook-mobile-design-margins.png)

### <a name="typography"></a><span data-ttu-id="e304c-175">Typographie</span><span class="sxs-lookup"><span data-stu-id="e304c-175">Typography</span></span>

<span data-ttu-id="e304c-176">La typographie est alignée sur Outlook iOS et doit être simple pour la lisibilité.</span><span class="sxs-lookup"><span data-stu-id="e304c-176">Typography usage is aligned to Outlook iOS and is kept simple for scannability.</span></span>

<span data-ttu-id="e304c-177">**Typographie sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-177">**Typography on iOS**</span></span>

![Exemples de typographie pour iOS](../images/outlook-mobile-design-typography.png)

<span data-ttu-id="e304c-179">**Typographie sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-179">**Typography on Android**</span></span>

![Exemples de typographie pour Android](../images/outlook-mobile-design-typography-android.png)

### <a name="color-palette"></a><span data-ttu-id="e304c-181">Palette de couleurs</span><span class="sxs-lookup"><span data-stu-id="e304c-181">Color palette</span></span>

<span data-ttu-id="e304c-p114">L’utilisation des couleurs est subtile dans Outlook iOS.  À des fins de cohérence, nous vous demandons d’utiliser les couleurs uniquement sur les actions et les erreurs, et que seule la barre de marque utilise une couleur unique.</span><span class="sxs-lookup"><span data-stu-id="e304c-p114">Color usage is subtle in Outlook iOS.  To align, we ask that usage of color is localized to actions and error states, with only the brand bar using a unique color.</span></span>

![Palette de couleurs pour iOS](../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a><span data-ttu-id="e304c-185">Cellules</span><span class="sxs-lookup"><span data-stu-id="e304c-185">Cells</span></span>

<span data-ttu-id="e304c-186">Étant donné que la barre de navigation ne peut pas être utilisée pour libeller une page, utilisez les titres de section pour libeller les pages.</span><span class="sxs-lookup"><span data-stu-id="e304c-186">Since the navigation bar cannot be used to label a page, use section titles to label pages.</span></span>

<span data-ttu-id="e304c-187">**Exemples de cellules sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-187">**Examples of cells on iOS**</span></span>

![Types de cellules pour iOS](../images/outlook-mobile-design-cell-types.png)
* * *
![Cellules « Do » pour iOS](../images/outlook-mobile-design-cell-dos.png)
* * *
![Cellules « Don’t » pour iOS](../images/outlook-mobile-design-cell-donts.png)
* * *
![Cellules et entrées pour iOS](../images/outlook-mobile-design-cell-input.png)

<span data-ttu-id="e304c-192">**Exemples de cellules sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-192">**Examples of cells on Android**</span></span>

![Types de cellules pour Android](../images/outlook-mobile-design-cell-type-android.png)
* * *
![Cellules « Do » pour Android](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![Cellules « Don’t » pour Android](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![Cellules et entrées pour Android, partie 1](../images/outlook-mobile-design-cell-input-1-android.png)

![Cellules et entrées pour Android, partie 2](../images/outlook-mobile-design-cell-input-2-android.png)

### <a name="actions"></a><span data-ttu-id="e304c-198">Actions</span><span class="sxs-lookup"><span data-stu-id="e304c-198">Actions</span></span>

<span data-ttu-id="e304c-199">Même si votre application gère une multitude d’actions, réfléchissez aux plus importantes que vous souhaitez que votre complément effectue, et concentrez-vous sur celles-ci.</span><span class="sxs-lookup"><span data-stu-id="e304c-199">Even if your app handles a multitude of actions, think about the most important ones you want your add-in to perform, and concentrate on those.</span></span>

<span data-ttu-id="e304c-200">**Exemples d’actions sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-200">**Examples of actions on iOS**</span></span>

![Actions et cellules dans iOS](../images/outlook-mobile-design-action-cells.png)
* * *
![Actions « Do » pour iOS](../images/outlook-mobile-design-action-dos.png)

<span data-ttu-id="e304c-203">**Exemples d’actions sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-203">**Examples of actions on Android**</span></span>

![Actions et cellules dans Android](../images/outlook-mobile-design-action-cells-android.png)
* * *
![Actions « Do » pour Android](../images/outlook-mobile-design-action-dos-android.png)

### <a name="buttons"></a><span data-ttu-id="e304c-206">Boutons</span><span class="sxs-lookup"><span data-stu-id="e304c-206">Buttons</span></span>

<span data-ttu-id="e304c-207">Les boutons sont utilisés lorsqu’il existe d’autres éléments de l’expérience utilisateur en dessous (par opposition aux actions, car une action est toujours le dernier élément de l’écran).</span><span class="sxs-lookup"><span data-stu-id="e304c-207">Buttons are used when there are other UX elements below (vs. actions, where the action is the last element on the screen).</span></span>

<span data-ttu-id="e304c-208">**Exemples de boutons sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-208">**Examples of buttons on iOS**</span></span>

![Exemples de boutons pour iOS](../images/outlook-mobile-design-buttons.png)

<span data-ttu-id="e304c-210">**Exemples de boutons sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-210">**Examples of buttons on Android**</span></span>

![Exemples de boutons pour Android](../images/outlook-mobile-design-buttons-android.png)

### <a name="tabs"></a><span data-ttu-id="e304c-212">Onglets</span><span class="sxs-lookup"><span data-stu-id="e304c-212">Tabs</span></span>

<span data-ttu-id="e304c-213">Les onglets peuvent contribuer à organiser le contenu.</span><span class="sxs-lookup"><span data-stu-id="e304c-213">Tabs can aid in content organization.</span></span>

<span data-ttu-id="e304c-214">**Exemples d’onglets sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-214">**Examples of tabs on iOS**</span></span>

![Exemples d’onglets pour iOS](../images/outlook-mobile-design-tabs.png)

<span data-ttu-id="e304c-216">**Exemples d’onglets sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-216">**Examples of tabs on Android**</span></span>

![Exemples d’onglets pour Android](../images/outlook-mobile-design-tabs-android.png)

### <a name="icons"></a><span data-ttu-id="e304c-218">Icônes</span><span class="sxs-lookup"><span data-stu-id="e304c-218">Icons</span></span>

<span data-ttu-id="e304c-p115">Les icônes doivent respecter la conception Outlook iOS actuelle autant que possible. Utilisez la taille et la couleur standard.</span><span class="sxs-lookup"><span data-stu-id="e304c-p115">Icons should follow the current Outlook iOS design when possible. Use our standard size and color.</span></span>

<span data-ttu-id="e304c-221">**Exemples d’icônes sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-221">**Examples of icons on iOS**</span></span>

![Exemples d’icônes pour iOS](../images/outlook-mobile-design-icons.png)

<span data-ttu-id="e304c-223">**Exemples d’icônes sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-223">**Examples of icons on Android**</span></span>

![Exemples d’icônes pour Android](../images/outlook-mobile-design-icons-android.jpg)

## <a name="end-to-end-examples"></a><span data-ttu-id="e304c-225">Exemples de bout en bout</span><span class="sxs-lookup"><span data-stu-id="e304c-225">End-to-end examples</span></span>

<span data-ttu-id="e304c-226">Pour le lancement de nos compléments Outlook Mobile v1, nous avons travaillé en étroite collaboration avec nos partenaires qui créaient des compléments. Pour présenter le potentiel de leurs compléments sur Outlook Mobile, notre concepteur a regroupé des flux de bout en bout pour chaque complément, en respectant nos instructions et en utilisant nos modèles.</span><span class="sxs-lookup"><span data-stu-id="e304c-226">For our v1 Outlook Mobile Add-ins launch, we worked closely with our partners who were building add-ins. As a way to showcase the potential of their add-ins on Outlook Mobile, our designer put together end-to-end flows for each add-in, leveraging our guidelines and patterns.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e304c-227">Ces exemples sont destinés à mettre en évidence la façon idéale de combiner interaction et conception visuelle pour un complément et peuvent ne pas correspondre aux ensembles de fonctionnalités exacts des compléments réels.</span><span class="sxs-lookup"><span data-stu-id="e304c-227">These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.</span></span> 

### <a name="giphy"></a><span data-ttu-id="e304c-228">GIPHY</span><span class="sxs-lookup"><span data-stu-id="e304c-228">GIPHY</span></span>

<span data-ttu-id="e304c-229">**Exemple de GIPHY sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-229">**An example of GIPHY on iOS**</span></span>

![Conception de bout en bout pour le complément GIPHY sur iOS](../images/outlook-mobile-design-giphy.png)

<span data-ttu-id="e304c-231">**Exemple de GIPHY sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-231">**An example of GIPHY on Android**</span></span>

![Conception de bout en bout pour le complément GIPHY sur Android](../images/outlook-mobile-design-giphy-android.png)

### <a name="nimble"></a><span data-ttu-id="e304c-233">Nimble</span><span class="sxs-lookup"><span data-stu-id="e304c-233">Nimble</span></span>

<span data-ttu-id="e304c-234">**Exemple de Nimble sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-234">**An example of Nimble on iOS**</span></span>

![Conception de bout en bout pour le complément Nimble sur iOS](../images/outlook-mobile-design-nimble.png)

<span data-ttu-id="e304c-236">**Exemple de Nimble sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-236">**An example of Nimble on Android**</span></span>

![Conception de bout en bout pour le complément Nimble sur Android](../images/outlook-mobile-design-nimble-android.png)

### <a name="trello"></a><span data-ttu-id="e304c-238">Trello</span><span class="sxs-lookup"><span data-stu-id="e304c-238">Trello</span></span>

<span data-ttu-id="e304c-239">**Exemple de Trello sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-239">**An example of Trello on iOS**</span></span>

![Conception de bout en bout pour le complément Trello partie 1 sur iOS](../images/outlook-mobile-design-trello-1.png)
* * *
![Conception de bout en bout pour le complément Trello partie 2 sur iOS](../images/outlook-mobile-design-trello-2.png)
* * *
![Conception de bout en bout pour le complément Trello partie 3 sur iOS](../images/outlook-mobile-design-trello-3.png)

<span data-ttu-id="e304c-243">**Exemple de Trello sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-243">**An example of Trello on Android**</span></span>

![Conception de bout en bout pour le complément Trello partie 1 sur Android](../images/outlook-mobile-design-trello-1-android.png)
* * *
![Conception de bout en bout pour le complément Trello partie 2 sur Android](../images/outlook-mobile-design-trello-2-android.png)

### <a name="dynamics-crm"></a><span data-ttu-id="e304c-246">Dynamics CRM</span><span class="sxs-lookup"><span data-stu-id="e304c-246">Dynamics CRM</span></span>

<span data-ttu-id="e304c-247">**Exemple de Dynamics CRM sur iOS**</span><span class="sxs-lookup"><span data-stu-id="e304c-247">**An example of Dynamics CRM on iOS**</span></span>

![Conception de bout en bout pour le complément Dynamics CRM sur iOS](../images/outlook-mobile-design-crm.png)

<span data-ttu-id="e304c-249">**Exemple de Dynamics CRM sur Android**</span><span class="sxs-lookup"><span data-stu-id="e304c-249">**An example of Dynamics CRM on Android**</span></span>

![Conception de bout en bout pour le complément Dynamics CRM sur Android](../images/outlook-mobile-design-crm-android.png)
