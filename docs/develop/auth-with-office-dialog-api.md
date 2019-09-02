---
title: Authentifier et autoriser avec l’API de boîte de dialogue Office
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 3d61c82f28fd5780176b356e1ab4d394e5fbf8bd
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302943"
---
# <a name="authenticate-and-authorize-with-the-office-dialog-api"></a><span data-ttu-id="51b53-102">Authentifier et autoriser avec l’API de boîte de dialogue Office</span><span class="sxs-lookup"><span data-stu-id="51b53-102">Authenticate and authorize with the Office Dialog API</span></span>

> [!NOTE]
> <span data-ttu-id="51b53-103">Cet article part du principe que vous avez l'habitude d’[utiliser l’API de boîte de dialogue](dialog-api-in-office-add-ins.md) dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="51b53-103">This article assumes that you are familiar with [Use the Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="51b53-104">De nombreuses autorités d’identité, également appelées service d’émission de jeton de sécurité (STS), empêchent leur page de connexion de s’ouvrir dans un IFRAME.</span><span class="sxs-lookup"><span data-stu-id="51b53-104">Many identity authorities, also called Secure Token Services (STS), prevent their login page from opening in an iframe.</span></span> <span data-ttu-id="51b53-105">Celles-ci incluent Google, Facebook et les services protégés par la plateforme d’identité Microsoft (anciennement Azure AD V 2.0) telles que le compte Microsoft et Office 365 (comptes professionnels ou scolaires).</span><span class="sxs-lookup"><span data-stu-id="51b53-105">These include Google, Facebook, and services protected by Microsoft Identity Platform (formerly Azure AD V 2.0) such as Microsoft Account and Office 365 (Work or School accounts).</span></span> <span data-ttu-id="51b53-106">Cela a pour effet de créer un problème pour les compléments Office, car lorsque le complément est exécuté dans **Office sur le Web**, le volet Office est un IFRAME.</span><span class="sxs-lookup"><span data-stu-id="51b53-106">This creates a problem for Office Add-ins because when the add-in is running in **Office on the web**, the task pane is an iframe.</span></span> <span data-ttu-id="51b53-107">Les utilisateurs d’un complément peuvent se connecter à l’un de ces services uniquement si le complément peut ouvrir une instance de navigateur entièrement distincte.</span><span class="sxs-lookup"><span data-stu-id="51b53-107">Users of an add-in can only login to one of these services if the add-in can open an entirely separate browser instance.</span></span> <span data-ttu-id="51b53-108">C’est la raison pour laquelle Office fournit son [API de boîte](dialog-api-in-office-add-ins.md) de dialogue, spécifiquement la méthode [displayDialogAsync](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="51b53-108">This is why Office provides its [Dialog API](dialog-api-in-office-add-ins.md), specifically the [displayDialogAsync](/javascript/api/office/office.ui) method.</span></span>

<span data-ttu-id="51b53-109">La boîte de dialogue ouverte avec cette API présente les caractéristiques suivantes :</span><span class="sxs-lookup"><span data-stu-id="51b53-109">The dialog box that is opened with this API has the following characteristics:</span></span>

- <span data-ttu-id="51b53-110">Elle n' [est pas modale](https://en.wikipedia.org/wiki/Dialog_box).</span><span class="sxs-lookup"><span data-stu-id="51b53-110">It is [non-modal](https://en.wikipedia.org/wiki/Dialog_box).</span></span>
- <span data-ttu-id="51b53-111">Il s’agit d’une instance de navigateur totalement distincte du volet de tâches, ce qui signifie :</span><span class="sxs-lookup"><span data-stu-id="51b53-111">It is a completely separate browser instance from the task pane, meaning:</span></span>
  - <span data-ttu-id="51b53-112">Elle possède ses propres environnements d’exécution JavaScript et objets de fenêtre et variables globales.</span><span class="sxs-lookup"><span data-stu-id="51b53-112">It has its own JavaScript runtime environment and window object and global variables.</span></span>
  - <span data-ttu-id="51b53-113">Il n’existe pas d’environnement d’exécution partagé dans le volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="51b53-113">There is no shared execution environment with the task pane.</span></span>
  - <span data-ttu-id="51b53-114">Elle ne partage pas le même espace de stockage de session que le volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="51b53-114">It does not share the same session storage as the task pane.</span></span>
- <span data-ttu-id="51b53-115">La première page ouverte dans la boîte de dialogue doit être hébergée dans le même domaine que le volet des tâches, y compris le protocole, les sous-domaines et le port, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="51b53-115">The first page opened in the dialog box must be hosted in the same domain as the task pane, including protocol, subdomains, and port, if any.</span></span>
- <span data-ttu-id="51b53-116">La boîte de dialogue peut renvoyer les informations au volet des tâches à l’aide de la méthode [messageParent](/javascript/api/office/office.ui#messageparent-message-), mais cette méthode ne peut être appelée que depuis une page hébergée dans le même domaine que le volet des tâches, y compris le protocole, les sous-domaines et le port.</span><span class="sxs-lookup"><span data-stu-id="51b53-116">The dialog box can send information back to the task pane by using the [messageParent](/javascript/api/office/office.ui#messageparent-message-) method, but this method can be called only from a page that is hosted in the same domain as the task pane, including protocol, subdomains, and port.</span></span>

<span data-ttu-id="51b53-117">Lorsque la boîte de dialogue n’est pas un IFRAME (qui est la valeur par défaut), elle peut ouvrir la page de connexion d’un fournisseur d’identité.</span><span class="sxs-lookup"><span data-stu-id="51b53-117">When the dialog is not an iframe (which is the default), it can open the login page of an identity provider.</span></span> <span data-ttu-id="51b53-118">Comme vous le verrez dans la section ci-dessous, les caractéristiques de la boîte de dialogue ont une incidence sur la manière dont vous utilisez les bibliothèques d’authentification ou d’autorisation telles que MSAL et Passport.</span><span class="sxs-lookup"><span data-stu-id="51b53-118">As you'll see below, the characteristics of the Dialog have implications for how you use authentication or authorization libraries such as MSAL and Passport.</span></span>

> [!NOTE]
> <span data-ttu-id="51b53-119">Vous pouvez configurer la boîte de dialogue pour qu’elle s’ouvre dans un IFRAME flottant : vous pouvez simplement transmettre l’option `displayInIframe: true`dans l’appel à`displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="51b53-119">There is a way to configure the dialog to open in a floating iframe: you simply pass the `displayInIframe: true` option in the call to `displayDialogAsync`.</span></span> <span data-ttu-id="51b53-120">Ne le faites *pas* lorsque vous utilisez l’API de boîte de dialogue pour la connexion.</span><span class="sxs-lookup"><span data-stu-id="51b53-120">Do *not* do this when you are using the Dialog API for login.</span></span>

## <a name="authentication-flow-with-the-dialog"></a><span data-ttu-id="51b53-121">Flux d’authentification avec la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="51b53-121">Authentication flow with the Dialog</span></span>

<span data-ttu-id="51b53-122">Voici un flux d’authentification simple et standard.</span><span class="sxs-lookup"><span data-stu-id="51b53-122">The following is a simple and typical authentication flow.</span></span> <span data-ttu-id="51b53-123">Les détails sont répertoriés après le diagramme.</span><span class="sxs-lookup"><span data-stu-id="51b53-123">Details are after the diagram.</span></span>

![Image illustrant la relation entre les processus du volet des tâches et du navigateur de boîte de dialogue.](../images/taskpane-dialog-processes.gif)

1. <span data-ttu-id="51b53-125">La première page qui s’ouvre dans la boîte de dialogue est une page (ou toute autre ressource) qui est hébergée dans le domaine du complément ; autrement dit, le même domaine que la fenêtre du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="51b53-125">The first page that opens in the dialog box is a page (or other resource) that is hosted in the add-in's domain; that is, the same domain as the task pane window.</span></span> <span data-ttu-id="51b53-126">Cette page peut avoir une IU simple indiquant « Veuillez patienter, nous allons vous rediriger vers la page sur laquelle vous pouvez vous connecter à *NOM DU FOURNISSEUR* ».</span><span class="sxs-lookup"><span data-stu-id="51b53-126">This page can have a simple UI that says "Please wait, we are redirecting you to the page where you can sign in to *NAME-OF-PROVIDER*."</span></span> <span data-ttu-id="51b53-127">Le code dans cette page construit l’URL de la page de connexion du fournisseur d’identité en utilisant les informations transmises à la boîte de dialogue, comme décrit dans [Transmission d’informations à la boîte de dialogue](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) ou est codée en dur dans un fichier de configuration du complément, tel qu’un fichier web.config.</span><span class="sxs-lookup"><span data-stu-id="51b53-127">The code in this page constructs the URL of the identity provider's sign-in page with information that is either passed to the dialog box as described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) or is hardcoded into a configuration file of the add-in, such as a web.config file.</span></span>
2. <span data-ttu-id="51b53-128">La fenêtre de dialogue redirige alors l’utilisateur vers la page de connexion.</span><span class="sxs-lookup"><span data-stu-id="51b53-128">The dialog window then redirects to the sign-in page.</span></span> <span data-ttu-id="51b53-129">L’URL inclut un paramètre de requête qui indique au fournisseur d’identité de rediriger la fenêtre de dialogue une fois que l’utilisateur s’est connecté à une page spécifique.</span><span class="sxs-lookup"><span data-stu-id="51b53-129">The URL includes a query parameter that tells the identity provider to redirect the dialog window, after the user signs in, to a specific page.</span></span> <span data-ttu-id="51b53-130">Dans cet article, nous appellerons cette page **redirectPage.html**.</span><span class="sxs-lookup"><span data-stu-id="51b53-130">In this article, we'll call this page "redirectPage.html".</span></span> <span data-ttu-id="51b53-131">*Il doit s’agir d’une page se trouvant dans le même domaine que la fenêtre hôte*, afin que les résultats de la tentative de connexion puissent être transférés au volet des tâches avec un appel de`messageParent`.</span><span class="sxs-lookup"><span data-stu-id="51b53-131">*This must be a page in the same domain as the host window*, so that the results of the sign-in attempt can be passed to the task pane with a call of `messageParent`.</span></span>
3. <span data-ttu-id="51b53-132">Le service du fournisseur d’identité traite la requête GET entrante à partir de la fenêtre de dialogue.</span><span class="sxs-lookup"><span data-stu-id="51b53-132">The identity provider's service processes the incoming GET request from the dialog window.</span></span> <span data-ttu-id="51b53-133">Si l’utilisateur est déjà connecté, il redirige immédiatement la fenêtre vers**redirectPage.html** et inclut les données utilisateur sous la forme d’un paramètre de requête.</span><span class="sxs-lookup"><span data-stu-id="51b53-133">If the user is already logged on, it immediately redirects the window to redirectPage.html and includes user data as a query parameter.</span></span> <span data-ttu-id="51b53-134">Si l’utilisateur n’est pas encore connecté, la page de connexion du fournisseur apparaît dans la fenêtre et l’utilisateur se connecte.</span><span class="sxs-lookup"><span data-stu-id="51b53-134">If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in.</span></span> <span data-ttu-id="51b53-135">Pour la plupart des fournisseurs, si l’utilisateur ne parvient pas à se connecter, le fournisseur affiche une page d’erreur dans la fenêtre de dialogue et ne redirige pas vers**redirectPage.html**.</span><span class="sxs-lookup"><span data-stu-id="51b53-135">For most providers, if the user cannot sign in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html.</span></span> <span data-ttu-id="51b53-136">L’utilisateur doit fermer la fenêtre en sélectionnant le **X** dans le coin.</span><span class="sxs-lookup"><span data-stu-id="51b53-136">The user must close the window by selecting the **X** in the corner.</span></span> <span data-ttu-id="51b53-137">Si l’utilisateur se connecte avec succès, la fenêtre de dialogue est redirigée vers**redirectPage.html** et les données utilisateur sont incluses sous la forme d’un paramètre de requête.</span><span class="sxs-lookup"><span data-stu-id="51b53-137">If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.</span></span>
4. <span data-ttu-id="51b53-138">Lorsque la page **redirectPage.html** s’ouvre, elle appelle`messageParent` pour indiquer le succès ou l’échec au volet des tâches et éventuellement indiquer également des données utilisateur ou des données d’erreur.</span><span class="sxs-lookup"><span data-stu-id="51b53-138">When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data.</span></span> <span data-ttu-id="51b53-139">Les autres messages possibles incluent le passage d’un jeton d’accès ou le volet des tâches dans lequel le jeton est stocké.</span><span class="sxs-lookup"><span data-stu-id="51b53-139">Other possible messages include passing an access token or telling the task pane that the token is in storage.</span></span>
5. <span data-ttu-id="51b53-140">L’événement `DialogMessageReceived` se déclenche dans le volet des tâches, et son gestionnaire ferme la fenêtre de dialogue et effectue éventuellement d’autres traitements du message.</span><span class="sxs-lookup"><span data-stu-id="51b53-140">The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message.</span></span>

#### <a name="support-multiple-identity-providers"></a><span data-ttu-id="51b53-141">Prise en charge de plusieurs fournisseurs d’identité</span><span class="sxs-lookup"><span data-stu-id="51b53-141">Support multiple identity providers</span></span>

<span data-ttu-id="51b53-p109">Si votre complément offre à l’utilisateur le choix entre plusieurs fournisseurs, tels qu’un compte Microsoft, Google ou Facebook, vous avez besoin d’une première page locale (voir section précédente) qui fournit une IU permettant à l’utilisateur de sélectionner un fournisseur. La sélection déclenche la construction de l’URL de connexion et la redirection vers celle-ci.</span><span class="sxs-lookup"><span data-stu-id="51b53-p109">If your add-in gives the user a choice of providers, such as Microsoft Account, Google, or Facebook, you need a local first page (see preceding section) that provides a UI for the user to select a provider. Selection triggers the construction of the sign-in URL and redirection to it.</span></span>

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a><span data-ttu-id="51b53-144">Autorisation du complément pour une ressource externe</span><span class="sxs-lookup"><span data-stu-id="51b53-144">Authorization of the add-in to an external resource</span></span>

<span data-ttu-id="51b53-145">Sur le web nouvelle génération, les applications web sont des principaux de sécurité au même titre que les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="51b53-145">In the modern web, users and web applications are security principals.</span></span> <span data-ttu-id="51b53-146">L’application a sa propre identité et ses propres autorisations pour une ressource en ligne comme Office 365, Google Plus, Facebook ou LinkedIn.</span><span class="sxs-lookup"><span data-stu-id="51b53-146">In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, or LinkedIn.</span></span> <span data-ttu-id="51b53-147">L’application est inscrite auprès du fournisseur de ressources avant d’être déployée.</span><span class="sxs-lookup"><span data-stu-id="51b53-147">The application is registered with the resource provider before it is deployed.</span></span> <span data-ttu-id="51b53-148">L’inscription inclut :</span><span class="sxs-lookup"><span data-stu-id="51b53-148">The registration includes:</span></span>

- <span data-ttu-id="51b53-149">La liste des autorisations dont l’application a besoin.</span><span class="sxs-lookup"><span data-stu-id="51b53-149">A list of the permissions that the application needs to a user's resources.</span></span>
- <span data-ttu-id="51b53-150">l’URL à laquelle le service de ressources doit renvoyer un jeton d’accès lorsque l’application accède au service.</span><span class="sxs-lookup"><span data-stu-id="51b53-150">A URL to which the resource service should return an access token when the application accesses the service.</span></span>  

<span data-ttu-id="51b53-p111">Lorsqu’un utilisateur appelle une fonction dans l’application qui accède aux données de l’utilisateur dans le service de ressources, l’utilisateur est invité à se connecter au service, puis à accorder à l’application les autorisations dont elle a besoin pour les ressources de l’utilisateur. Ensuite, le service redirige la fenêtre de connexion vers l’URL précédemment inscrite et transmet le jeton d’accès. L’application utilise le jeton d’accès pour accéder aux ressources de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="51b53-p111">When a user invokes a function in the application that accesses the user's data in the resource service, they are prompted to sign in to the service and then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources.</span></span>

<span data-ttu-id="51b53-154">Vous pouvez utiliser les API de dialogue pour gérer ce processus à l’aide d’un flux semblable à celui décrit pour la connexion des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="51b53-154">You can use the Dialog APIs to manage this process by using a flow that is similar to the one described for users to sign in.</span></span> <span data-ttu-id="51b53-155">Les seules différences sont les suivantes :</span><span class="sxs-lookup"><span data-stu-id="51b53-155">The only differences are:</span></span>

- <span data-ttu-id="51b53-156">Si l’utilisateur n’a pas préalablement accordé à l’application les autorisations nécessaires, il est invité à le faire dans la boîte de dialogue après la connexion.</span><span class="sxs-lookup"><span data-stu-id="51b53-156">If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog box after signing in.</span></span>
- <span data-ttu-id="51b53-157">La fenêtre de dialogue envoie le jeton d’accès à la fenêtre hôte en utilisant `messageParent` pour envoyer le jeton d’accès converti en chaîne ou en stockant jeton d’accès à un emplacement où la fenêtre hôte peut le récupérer (et utilise `messageParent` pour indiquer à la fenêtre hôte que le jeton est disponible).</span><span class="sxs-lookup"><span data-stu-id="51b53-157">The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it.</span></span> <span data-ttu-id="51b53-158">Le jeton a une limite de temps, mais tant qu’elle n’est pas écoulée, la fenêtre hôte peut l’utiliser pour accéder directement aux ressources de l’utilisateur sans demander d’autre confirmation.</span><span class="sxs-lookup"><span data-stu-id="51b53-158">The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting.</span></span>

<span data-ttu-id="51b53-159">Quelques exemples de compléments d’authentification qui utilisent l’API de boîte de dialogue à cet effet sont répertoriés dans les [exemples](#samples).</span><span class="sxs-lookup"><span data-stu-id="51b53-159">Some authentication sample add-ins that use the Dialog API for this purpose are listed in [Samples](#samples).</span></span>

## <a name="using-authentication-libraries-with-the-dialog"></a><span data-ttu-id="51b53-160">Utilisation de bibliothèques d’authentification avec la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="51b53-160">Using authentication libraries with the dialog</span></span>

<span data-ttu-id="51b53-161">Le fait que la boîte de dialogue Office et le volet des tâches s’exécutent dans différents navigateurs, et instances JavaScript Runtime, signifie que vous devez utiliser de nombreuses bibliothèques d’authentification et d’autorisation de manière différente que celle utilisée lorsque l’authentification et l’autorisation peuvent être effectuées dans la même fenêtre.</span><span class="sxs-lookup"><span data-stu-id="51b53-161">The fact that the Office Dialog and the task pane run in different browser, and JavaScript runtime, instances means that you must use many authentication/authorization libraries in the way that is different from how they are used when authentication and authorization can take place in the same window.</span></span> <span data-ttu-id="51b53-162">Les sections suivantes décrivent les principales façons dont vous ne pouvez généralement pas utiliser ces bibliothèques et la *manière*de les utiliser.</span><span class="sxs-lookup"><span data-stu-id="51b53-162">The following sections describe the main ways that you usually cannot use these libraries and the way that you *can* use them.</span></span>

### <a name="you-usually-cannot-use-the-librarys-internal-cache-to-store-tokens"></a><span data-ttu-id="51b53-163">En général, vous ne pouvez pas utiliser le cache interne de la bibliothèque pour stocker des jetons</span><span class="sxs-lookup"><span data-stu-id="51b53-163">You usually cannot use the library's internal cache to store tokens</span></span>

<span data-ttu-id="51b53-164">En règle générale, les bibliothèques associées à l’authentification fournissent un cache en mémoire pour stocker le jeton d’accès.</span><span class="sxs-lookup"><span data-stu-id="51b53-164">Typically, auth-related libraries provide an in-memory cache to store the access token.</span></span> <span data-ttu-id="51b53-165">Si des appels ultérieurs au fournisseur de ressources (par exemple, Google, Microsoft Graph, Facebook, etc.) sont apportés, la bibliothèque vérifie tout d’abord si le jeton dans son cache a expiré.</span><span class="sxs-lookup"><span data-stu-id="51b53-165">If subsequent calls to the resource provider (such as Google, Microsoft Graph, Facebook, etc.) are made, the library will first check to see if the token in its cache is expired.</span></span> <span data-ttu-id="51b53-166">Si celui-ci n’a pas expiré, la bibliothèque renvoie le jeton mis en cache plutôt que d’effectuer un autre aller-retour vers le SJS pour un nouveau jeton.</span><span class="sxs-lookup"><span data-stu-id="51b53-166">If it is unexpired, the library returns the cached token rather than making another round-trip to the STS for a new token.</span></span> <span data-ttu-id="51b53-167">Mais ce modèle n’est pas utilisable dans les compléments Office. Dans la mesure où la connexion a lieu dans l’instance de navigateur de la boîte de dialogue Office, le cache de jetons est dans cette instance.</span><span class="sxs-lookup"><span data-stu-id="51b53-167">But this pattern is not usable in Office add-ins. Since the login occurs in the Office Dialog's browser instance, the token cache is in that instance.</span></span>

<span data-ttu-id="51b53-168">Ceci est lié au fait qu’une bibliothèque fournit généralement des méthodes à la fois interactives et «silencieuses» pour obtenir un jeton.</span><span class="sxs-lookup"><span data-stu-id="51b53-168">Closely related to this is the fact that a library will typically provide both interactive and "silent" methods for getting a token.</span></span> <span data-ttu-id="51b53-169">Lorsque vous pouvez effectuer les deux appels d’authentification et de données à la ressource dans la même instance de navigateur, votre code appelle la méthode silencieuse pour obtenir un jeton juste avant que votre code n’ajoute le jeton à l’appel de données.</span><span class="sxs-lookup"><span data-stu-id="51b53-169">When you can do both the authentication and the data calls to the resource in the same browser instance, your code calls the silent method to obtain a token just before your code adds the token to the data call.</span></span> <span data-ttu-id="51b53-170">La méthode silencieuse vérifie la présence d’un jeton non expiré dans le cache et le renvoie, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="51b53-170">The silent method checks for an unexpired token in the cache and returns it, if there is one.</span></span> <span data-ttu-id="51b53-171">Dans le cas contraire, la méthode silencieuse appelle la méthode interactive qui redirige vers la connexion de STS.</span><span class="sxs-lookup"><span data-stu-id="51b53-171">Otherwise, the silent method calls the interactive method which redirects to the STS's login.</span></span> <span data-ttu-id="51b53-172">Une fois la connexion terminée, la méthode interactive renvoie le jeton, mais le met en cache dans la mémoire.</span><span class="sxs-lookup"><span data-stu-id="51b53-172">After login completes, the interactive method returns the token, but also caches it in memory.</span></span> <span data-ttu-id="51b53-173">En revanche, lorsque l’API de boîte de dialogue Office est utilisée, les données appellent la ressource, qui appellent la méthode silencieuse, se trouvent dans l’instance de navigateur du volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="51b53-173">But when the Office Dialog API is being used, the data calls to the resource, which would call the silent method, are in the task pane's browser instance.</span></span> <span data-ttu-id="51b53-174">Le cache de jetons de la bibliothèque n’existe pas dans cette instance.</span><span class="sxs-lookup"><span data-stu-id="51b53-174">The library's token cache does not exist in that instance.</span></span>

<span data-ttu-id="51b53-175">En guise d’alternative, l’instance de navigateur de la boîte de dialogue de votre complément peut appeler directement la méthode interactive de la bibliothèque.</span><span class="sxs-lookup"><span data-stu-id="51b53-175">As an alternative, your add-in's Dialog browser instance can directly call the library's interactive method.</span></span> <span data-ttu-id="51b53-176">Lorsque cette méthode renvoie un jeton, votre code doit stocker de manière explicite le jeton à l’endroit où l’instance de navigateur du volet des tâches peut le récupérer (par exemple, stockage local ou une base de données côté serveur).</span><span class="sxs-lookup"><span data-stu-id="51b53-176">When that method returns a token, your code must explicitly store the token someplace where the task pane's browser instance can retrieve it, such as Local Storage or a server-side database.</span></span> <span data-ttu-id="51b53-177">Une autre option consiste à transmettre le jeton au volet des tâches avec la méthode`messageParent`.</span><span class="sxs-lookup"><span data-stu-id="51b53-177">Another option is to pass the token to the task pane with the `messageParent` method.</span></span> <span data-ttu-id="51b53-178">Cette alternative est uniquement possible si la méthode interactive stocke le jeton d’accès à un endroit où votre code peut le lire.</span><span class="sxs-lookup"><span data-stu-id="51b53-178">This alternative is only possible if the interactive method stores the access token in a place where your code can read it.</span></span> <span data-ttu-id="51b53-179">Parfois, la méthode interactive d’une bibliothèque est conçue pour stocker le jeton dans une propriété privée d’un objet qui n’est pas accessible à votre code.</span><span class="sxs-lookup"><span data-stu-id="51b53-179">Sometimes a library's interactive method is designed to store the token in a private property of an object that is inaccessible to your code.</span></span>

### <a name="you-usually-cannot-use-the-librarys-auth-context-object"></a><span data-ttu-id="51b53-180">En général, vous ne pouvez pas utiliser l’objet «contexte d’authentification» de la bibliothèque.</span><span class="sxs-lookup"><span data-stu-id="51b53-180">You usually cannot use the library's "auth context" object</span></span>

<span data-ttu-id="51b53-181">Il arrive souvent qu’une bibliothèque liée à l’authentification ait une méthode qui récupère un jeton de façon interactive et crée également un objet «contexte d’authentification» que la méthode renvoie.</span><span class="sxs-lookup"><span data-stu-id="51b53-181">Often, an auth-related library has a method that both obtains a token interactively and also creates an "auth-context" object which the method returns.</span></span> <span data-ttu-id="51b53-182">Le jeton est une propriété de l’objet (potentiellement privé et inaccessible directement à partir de votre code).</span><span class="sxs-lookup"><span data-stu-id="51b53-182">The token is a property of the object (possibly private and inaccessible directly from your code).</span></span> <span data-ttu-id="51b53-183">Cet objet possède les méthodes pour recevoir les données de la ressource.</span><span class="sxs-lookup"><span data-stu-id="51b53-183">That object has the methods that get data from the resource.</span></span> <span data-ttu-id="51b53-184">Ces méthodes incluent le jeton dans les requêtes HTTP qu’ils font au fournisseur de ressources (par exemple, Google, Microsoft Graph, Facebook, etc.).</span><span class="sxs-lookup"><span data-stu-id="51b53-184">These methods include the token in the HTTP Requests that they make to the resource provider (such as Google, Microsoft Graph, Facebook, etc.).</span></span>

<span data-ttu-id="51b53-185">Ces objets de contexte d’authentification, ainsi que les méthodes qui les créent, ne sont pas utilisables dans les compléments Office. Dans la mesure où la connexion a lieu dans l’instance de navigateur de la boîte de dialogue Office, l’objet doit être créé à cet emplacement.</span><span class="sxs-lookup"><span data-stu-id="51b53-185">These auth-context objects, and the methods that create them, are not usable in Office add-ins. Since the login occurs in the Office Dialog's browser instance, the object would have to be created there.</span></span> <span data-ttu-id="51b53-186">Mais les appels de données à la ressource se trouvent dans l’instance de navigateur du volet des tâches et il n’est pas possible d’utiliser l’objet d’une instance à l’autre.</span><span class="sxs-lookup"><span data-stu-id="51b53-186">But the data calls to the resource are in the task pane browser instance and there is no way to get the object from one instance to another.</span></span> <span data-ttu-id="51b53-187">Par exemple, vous ne pouvez pas passer l'objet avec`messageParent` car `messageParent`peut uniquement transmettre des chaînes ou des valeurs booléennes.</span><span class="sxs-lookup"><span data-stu-id="51b53-187">For example, you cannot pass the object with `messageParent` because `messageParent` can only pass strings or boolean values.</span></span> <span data-ttu-id="51b53-188">Un objet JavaScript avec des méthodes ne peut pas être mis en chaîne de façon fiable.</span><span class="sxs-lookup"><span data-stu-id="51b53-188">A JavaScript object with methods cannot be reliably stringified.</span></span>

### <a name="how-you-can-use-libraries-with-the-office-dialog-api"></a><span data-ttu-id="51b53-189">Utilisation des bibliothèques avec l’API de boîte de dialogue Office</span><span class="sxs-lookup"><span data-stu-id="51b53-189">How you can use libraries with the Office Dialog API</span></span>

<span data-ttu-id="51b53-190">En plus ou au lieu de, des objets «contexte d’authentification» monolithiques, la plupart des bibliothèques fournissent des API à un niveau d’abstraction inférieur qui permettent à votre code de créer moins d’objets d’assistance monolithiques.</span><span class="sxs-lookup"><span data-stu-id="51b53-190">In addition to, or instead of, monolithic "auth context" objects, most libraries provide APIs at a lower level of abstraction that enable your code to create less monolithic helper objects.</span></span> <span data-ttu-id="51b53-191">Par exemple, [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) v.</span><span class="sxs-lookup"><span data-stu-id="51b53-191">For example, [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) v.</span></span> <span data-ttu-id="51b53-192">3. x. x a une API pour créer une URL de connexion, et une autre API qui crée un objet AuthResult qui contient un jeton d’accès dans une propriété accessible à votre code.</span><span class="sxs-lookup"><span data-stu-id="51b53-192">3.x.x has an API to construct a login URL, and another API that constructs an AuthResult object that contains an access token in a property that is accessible to your code.</span></span> <span data-ttu-id="51b53-193">Pour consulter des exemples d’MSAL.net dans un complément Office, voir :[complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) et [complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET).</span><span class="sxs-lookup"><span data-stu-id="51b53-193">For examples of MSAL.NET in an Office add-in see: [Office Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) and [Outlook Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET).</span></span>

<span data-ttu-id="51b53-194">Pour plus d’informations sur les bibliothèques d’authentification et d’autorisation, voir [Microsoft Graph : bibliothèques recommandées](authorize-to-microsoft-graph-without-sso.md#recommended-libraries-and-samples) et [autres services externes : bibliothèques](auth-external-add-ins.md#libraries).</span><span class="sxs-lookup"><span data-stu-id="51b53-194">For more information about authentication and authorization libraries, see [Microsoft Graph: Recommended libraries](authorize-to-microsoft-graph-without-sso.md#recommended-libraries-and-samples) and [Other external services: Libraries](auth-external-add-ins.md#libraries).</span></span>

## <a name="samples"></a><span data-ttu-id="51b53-195">Exemples</span><span class="sxs-lookup"><span data-stu-id="51b53-195">Samples</span></span>

- <span data-ttu-id="51b53-196">[Complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET): complément ASP.net (Excel, Word ou PowerPoint) qui utilise la bibliothèque MSAL.net pour se connecter et obtenir un jeton d’accès pour les données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="51b53-196">[Office Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET): An ASP.NET based add-in (Excel, Word, or PowerPoint) that uses the MSAL.NET library to login and get an access token for Microsoft Graph data.</span></span>
- <span data-ttu-id="51b53-197">[Complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET): comme celui ci-dessus, mais l’application Office est Outlook.</span><span class="sxs-lookup"><span data-stu-id="51b53-197">[Outlook Add-in Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET): Just like the one above, but the Office application is Outlook.</span></span>

<span data-ttu-id="51b53-198">Pour plus d’informations, voir :</span><span class="sxs-lookup"><span data-stu-id="51b53-198">For more information, see:</span></span>
- [<span data-ttu-id="51b53-199">Autoriser des services externes dans votre complément Office</span><span class="sxs-lookup"><span data-stu-id="51b53-199">Authorize external services in your Office Add-in</span></span>](auth-external-add-ins.md)
- [<span data-ttu-id="51b53-200">Utiliser l’API de dialogue dans vos compléments Office</span><span class="sxs-lookup"><span data-stu-id="51b53-200">Use the Dialog API in your Office Add-ins</span></span>](dialog-api-in-office-add-ins.md)
