# <a name="navigation-patterns"></a><span data-ttu-id="78df8-101">Modèles de navigation</span><span class="sxs-lookup"><span data-stu-id="78df8-101">Navigation patterns</span></span>

<span data-ttu-id="78df8-102">Les principales fonctionnalités d’un complément sont accessibles via des types de commandes spécifiques et une zone d’écran limitée.</span><span class="sxs-lookup"><span data-stu-id="78df8-102">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="78df8-103">Il est important que la navigation soit intuitive, qu'elle fournisse un contexte et permette à l‘utilisateur de se déplacer facilement dans toute l'étendue du complément.</span><span class="sxs-lookup"><span data-stu-id="78df8-103">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="78df8-104">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="78df8-104">Best practices</span></span>

| <span data-ttu-id="78df8-105">À faire</span><span class="sxs-lookup"><span data-stu-id="78df8-105">Do</span></span>    | <span data-ttu-id="78df8-106">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="78df8-106">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="78df8-107">Assurez-vous que l’utilisateur dispose d’une option de navigation clairement visible.</span><span class="sxs-lookup"><span data-stu-id="78df8-107">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="78df8-108">Ne compliquez pas le processus de navigation en utilisant une interface utilisateur non standard.</span><span class="sxs-lookup"><span data-stu-id="78df8-108">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="78df8-109">Utilisez les composants suivants, le cas échéant, pour permettre aux utilisateurs de naviguer sur l'étendue de votre complément.</span><span class="sxs-lookup"><span data-stu-id="78df8-109">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="78df8-110">L’utilisateur doit comprendre sa place ou son contexte actuel dans le complément, c’est pourquoi vous ne devez pas lui compliquer la tâche</span><span class="sxs-lookup"><span data-stu-id="78df8-110">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="78df8-111">Barre de commandes</span><span class="sxs-lookup"><span data-stu-id="78df8-111">command bar</span></span>

<span data-ttu-id="78df8-112">La barre de commandes est une surface qui héberge des commandes qui agissent sur le contenu de la fenêtre, du panneau ou de la région parent au-dessus de laquelle elle réside.</span><span class="sxs-lookup"><span data-stu-id="78df8-112">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="78df8-113">Les caractéristiques facultatives comprennent le point d’accès au menu hamburger, la recherche et des commandes latérales.</span><span class="sxs-lookup"><span data-stu-id="78df8-113">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Commandes — Spécifications pour le volet des tâches de bureau](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="78df8-115">Barre d’onglets</span><span class="sxs-lookup"><span data-stu-id="78df8-115">Tab bar</span></span>

<span data-ttu-id="78df8-116">Barre d’onglets — Affiche la navigation à l’aide de boutons avec du texte et des icônes verticalement empilés.</span><span class="sxs-lookup"><span data-stu-id="78df8-116">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="78df8-117">Utiliser la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.</span><span class="sxs-lookup"><span data-stu-id="78df8-117">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Barre d’onglets — Spécifications du volet des tâches de bureau](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="78df8-119">Retour</span><span class="sxs-lookup"><span data-stu-id="78df8-119">Back button</span></span>

<span data-ttu-id="78df8-120">Le bouton Précédent permet aux utilisateurs de revenir au stade intial après avoir exécuté une action de navigation détaillée.</span><span class="sxs-lookup"><span data-stu-id="78df8-120">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="78df8-121">Utilisez ce modèle pour vous assurer que les utilisateurs suivent une série d’étapes ordonnées.</span><span class="sxs-lookup"><span data-stu-id="78df8-121">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![Bouton Précédent — Spécifications du volet des tâches de bureau](../images/add-in-back-button.png)
