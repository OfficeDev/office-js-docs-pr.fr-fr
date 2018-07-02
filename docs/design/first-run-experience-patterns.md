# <a name="first-run-experience-patterns"></a><span data-ttu-id="23b67-101">Modèles d'expérience de première exécution</span><span class="sxs-lookup"><span data-stu-id="23b67-101">First-run experience patterns</span></span>

<span data-ttu-id="23b67-102">Une première expérience d'exécution (FRE) est la première prise en main de votre complément par un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="23b67-102">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="23b67-103">Une FRE est présentée lorsqu'un utilisateur ouvre un complément pour la première fois et elle lui fournit un aperçu des fonctions, fonctionnalités ou avantages du complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-103">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="23b67-104">Cette expérience aide à façonner l'impression de l'utilisateur sur un complément et peut fortement l'influencer pour qu'il revienne et continue à utiliser votre complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-104">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="23b67-105">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="23b67-105">Best practices</span></span>


<span data-ttu-id="23b67-106">Suivez ces bonnes pratiques pour développer votre première expérience d'exécution :</span><span class="sxs-lookup"><span data-stu-id="23b67-106">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="23b67-107">À faire</span><span class="sxs-lookup"><span data-stu-id="23b67-107">Do</span></span>|<span data-ttu-id="23b67-108">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="23b67-108">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="23b67-109">Fournir une présentation simple et brève des principales actions du complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-109">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="23b67-110">Exclure les informations et les appels qui ne sont pas utiles pour commencer.</span><span class="sxs-lookup"><span data-stu-id="23b67-110">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="23b67-111">Donner aux utilisateurs l'opportunité d'effectuer une action qui aura un impact positif sur leur utilisation du complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-111">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="23b67-112">Ne pas vous attendre pas à ce que les utilisateurs apprennent tout à la fois.</span><span class="sxs-lookup"><span data-stu-id="23b67-112">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="23b67-113">Se concentrer sur l'action qui apporte le plus de valeur.</span><span class="sxs-lookup"><span data-stu-id="23b67-113">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="23b67-114">Créer une expérience attrayante que les utilisateurs voudront effectuer.</span><span class="sxs-lookup"><span data-stu-id="23b67-114">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="23b67-115">Ne pas obliger les utilisateurs à cliquer au cours de la première expérience d'exécution.</span><span class="sxs-lookup"><span data-stu-id="23b67-115">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="23b67-116">Donner aux utilisateurs une option pour contourner l'expérience de première exécution.</span><span class="sxs-lookup"><span data-stu-id="23b67-116">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="23b67-117">Déterminer s’il convient de montrer aux utilisateurs la première expérience d’utilisation une ou plusieurs fois (tout dépend de son importance pour votre scénario).</span><span class="sxs-lookup"><span data-stu-id="23b67-117">Consider whether showing users the first-run experience once or many times is important to your scenario.</span></span> <span data-ttu-id="23b67-118">Par exemple, si votre complément n'est utilisé que périodiquement, les utilisateurs peuvent devenir moins familiers et une nouvelle interaction avec l'expérience de première exécution peut être utile.</span><span class="sxs-lookup"><span data-stu-id="23b67-118">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="23b67-119">Appliquer les modèles suivants, le cas échéant, pour créer ou améliorer la première expérience d'exécution de votre complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-119">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="23b67-120">Carrousel</span><span class="sxs-lookup"><span data-stu-id="23b67-120">Carousel</span></span>


<span data-ttu-id="23b67-121">Le carrousel présente aux utilisateurs une série de fonctionnalités ou de pages d'informations avant qu'ils commencent à utiliser le complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-121">Walkthrough takes users through a series of features or information before they start using the add-in. (PDF, code)</span></span>

<span data-ttu-id="23b67-122">*Figure 1 : Autoriser les utilisateurs à dérouler ou à ignorer les premières pages du flux du carrousel.*
![Carrousel de première exécution - Spécifications pour le volet des tâches de bureau](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="23b67-122">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="23b67-123">*Figure 2 : Limiter au maximum le nombre d'écrans de carrousel afin de ne présenter à l'utilisateur que ce qui est nécessaire pour faire passer efficacement votre message*
![Carrousel de première exécution - Spécifications pour le volet des tâches de bureau](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="23b67-123">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="23b67-124">*Figure 3: Fournir un appel clair à l'action pour quitter l'expérience de première exécution.*
![Carrousel de première exécution - Spécifications pour le volet des tâches de bureau](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="23b67-124">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="23b67-125">Canevas de valeur</span><span class="sxs-lookup"><span data-stu-id="23b67-125">Value Placemat</span></span>

<span data-ttu-id="23b67-126">Le canevas de valeur illustre la proposition de valeur de votre complément par l'intermédiaire d'un logo, d'une proposition de valeur clairement énoncée, de faits saillants ou d'un résumé de caractéristiques, et d'une incitation à l'action.</span><span class="sxs-lookup"><span data-stu-id="23b67-126">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="23b67-127">![Première exécution - Canevas de valeur - Spécifications pour le volet des tâches de bureau](../images/add-in-FRE-value.png)
*Un canevas de valeur avec logo, proposition de valeur claire, résumé des fonctionnalités et appel à l'action.*</span><span class="sxs-lookup"><span data-stu-id="23b67-127">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="23b67-128">Canevas vidéo</span><span class="sxs-lookup"><span data-stu-id="23b67-128">Video Placemat</span></span>

<span data-ttu-id="23b67-129">Le canevas vidéo montre une vidéo aux utilisateurs avant qu’ils commencent à utiliser votre complément.</span><span class="sxs-lookup"><span data-stu-id="23b67-129">Video shows users a video before they start using your add-in. (spec, code)</span></span>


<span data-ttu-id="23b67-130">*Figure 1 : Canevas première exécution - L'écran contient une image fixe extraite de la vidéo avec un bouton de lecture et efface le bouton d'appel à l'action.*![Canevas vidéo - Spécifications pour le volet des tâches de bureau](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="23b67-130">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="23b67-131">*Figure 2 : Lecteur vidéo - Une vidéo est présentée aux utilisateurs dans une fenêtre de dialogue.*
![Canevas vidéo - Spécifications pour le volet des tâches de bureau](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="23b67-131">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
