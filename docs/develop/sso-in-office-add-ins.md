---
title: Activer l’authentification unique pour des compléments Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: dc9050d574e0a5e74ae8cae2c63817aa4f952eb9
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691194"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="4d551-102">Activer l’authentification unique pour des compléments Office (aperçu)</span><span class="sxs-lookup"><span data-stu-id="4d551-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="4d551-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span><span class="sxs-lookup"><span data-stu-id="4d551-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Image illustrant le processus de connexion pour un complément](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="4d551-106">État de Prévisualisation</span><span class="sxs-lookup"><span data-stu-id="4d551-106">Preview Status</span></span>

<span data-ttu-id="4d551-107">L’API de l’authentification unique est actuellement prise en charge uniquement en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="4d551-107">The Single Sign-on API is currently supported in preview only.</span></span> <span data-ttu-id="4d551-108">Elle est disponible pour les développeurs à des fins d’expérimentation ; mais elle ne doit pas être utilisée dans un complément de production.</span><span class="sxs-lookup"><span data-stu-id="4d551-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="4d551-109">Par ailleurs, les compléments qui utilisent l’authentification unique SSO ne sont pas acceptés dans [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="4d551-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="4d551-110">L’authentification unique requiert Office 365 (version d’Office par abonnement).</span><span class="sxs-lookup"><span data-stu-id="4d551-110">SSO requires Office 365 (the subscription version of Office, also called “Click to Run”).</span></span> <span data-ttu-id="4d551-111">Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider.</span><span class="sxs-lookup"><span data-stu-id="4d551-111">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="4d551-112">Vous devez participer au programme Office Insider pour obtenir cette version.</span><span class="sxs-lookup"><span data-stu-id="4d551-112">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="4d551-113">Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="4d551-113">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="4d551-114">Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.</span><span class="sxs-lookup"><span data-stu-id="4d551-114">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

<span data-ttu-id="4d551-115">Toutes les applications Office ne prennent pas en charge la prévisualisation de l’authentification unique (SSO).</span><span class="sxs-lookup"><span data-stu-id="4d551-115">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="4d551-116">Elle est disponible dans Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4d551-116">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="4d551-117">Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="4d551-117">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="4d551-118">Meilleures Pratiques et Conditions Requises</span><span class="sxs-lookup"><span data-stu-id="4d551-118">Requirements and Best Practices</span></span>

<span data-ttu-id="4d551-119">Pour utiliser l’authentification unique SSO, vous devez télécharger la version bêta de la bibliothèque JavaScript d’Office à partir de `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` dans la page de démarrage HTML du complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-119">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="4d551-120">Si vous utilisez un complément\*\* Outlook\*\*, veillez à activer l’Authentification Moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="4d551-120">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="4d551-121">Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="4d551-121">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="4d551-122">Vous ne devez*pas* dépendre de l’authentification unique SSO comme seule méthode de votre complément d’authentification.</span><span class="sxs-lookup"><span data-stu-id="4d551-122">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="4d551-123">Vous devez implémenter un système d’authentification secondaire vers lequel votre complément peut revenir dans certaines situations d’erreur.</span><span class="sxs-lookup"><span data-stu-id="4d551-123">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="4d551-124">Vous pouvez utiliser un système de tableaux d’utilisateur et d’authentification, ou vous pouvez tirer parti d’un des fournisseurs de connexion sociale.</span><span class="sxs-lookup"><span data-stu-id="4d551-124">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="4d551-125">Pour plus d’informations sur la procédure à suivre avec un complément Office, voir[Services externes autorisées dans votre complément Office](/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4d551-125">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="4d551-126">Concernant*Outlook*, il existe un système fall back recommandé.</span><span class="sxs-lookup"><span data-stu-id="4d551-126">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="4d551-127">Pour plus d’informations, voir[Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="4d551-127">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="4d551-128">Mode de fonctionnement de l’authentification unique SSO en cours d’exécution</span><span class="sxs-lookup"><span data-stu-id="4d551-128">How SSO works at runtime</span></span>

<span data-ttu-id="4d551-129">Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique SSO.</span><span class="sxs-lookup"><span data-stu-id="4d551-129">The following diagram shows how the SSO process works.</span></span>

![Un diagramme illustrant le processus d’authentification unique SSO](../images/sso-overview-diagram.png)

1. <span data-ttu-id="4d551-131">Dans le complément, JavaScript appelle une nouvelle API[getAccessTokenAsync](#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="4d551-131">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference).</span></span> <span data-ttu-id="4d551-132">Cela indique à l’application hôte Office qu’elle doit obtenir un jeton d’accès au complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-132">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="4d551-133">Voir [Exemple de token](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="4d551-133">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="4d551-134">Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.</span><span class="sxs-lookup"><span data-stu-id="4d551-134">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="4d551-135">Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.</span><span class="sxs-lookup"><span data-stu-id="4d551-135">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="4d551-136">L’application hôte Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="4d551-136">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="4d551-137">Azure AD envoie le jeton de complément à l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="4d551-137">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="4d551-138">L’application hôte Office envoie le**jeton de complément (token)** au complément dans le cadre de l’objet de résultat renvoyé par l’appel`getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="4d551-138">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="4d551-139">Dans le compl?ment, JavaScript peut analyser le token et extraire les informations dont il a besoin, telles que l'adresse e-mail de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d551-139">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="4d551-140">Optionnellement, le compl?ment peut envoyer une requ?te HTTP ? son serveur pour obtenir plus de donn?es sur l'utilisateur, notamment les pr?f?rences de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d551-140">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="4d551-141">Alternativement, le token lui-m?me pourrait ?tre envoy? au serveur pour analyse et validation.</span><span class="sxs-lookup"><span data-stu-id="4d551-141">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="4d551-142">Développer un complément d’authentification unique SSO</span><span class="sxs-lookup"><span data-stu-id="4d551-142">Develop an SSO add-in</span></span>

<span data-ttu-id="4d551-p109">Cette section décrit les tâches impliquées dans la création d’un complément Office qui utilise l’authentification unique. Ces tâches sont décrites ici indépendamment du langage et de l’infrastructure. Pour obtenir des exemples de procédures pas à pas détaillées, consultez les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="4d551-p109">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="4d551-146">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="4d551-146">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="4d551-147">Créer un complément Office ASP.NET qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="4d551-147">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="4d551-148">Créer l’application de service</span><span class="sxs-lookup"><span data-stu-id="4d551-148">Create the service application</span></span>

<span data-ttu-id="4d551-p110">Enregistrez le complément sur le portail d’inscription pour le point de terminaison Azure v2.0 : https://apps.dev.microsoft.com. Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :</span><span class="sxs-lookup"><span data-stu-id="4d551-p110">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="4d551-151">Obtenez un ID client et un code secret pour le complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-151">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="4d551-152">Spécifiez les autorisations dont votre complément a besoin pour AAD v.</span><span class="sxs-lookup"><span data-stu-id="4d551-152">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="4d551-153"> Point de terminaison 2.0 (et ?ventuellement Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="4d551-153">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="4d551-154">L'autorisation "profil" est toujours n?cessaire.</span><span class="sxs-lookup"><span data-stu-id="4d551-154">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="4d551-155">Accordez la confiance de l’application hôte Office au complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-155">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="4d551-156">Pré-autorisez l’application hôte Office pour le complément avec l’autorisation par défaut*access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="4d551-156">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="4d551-157">Pour plus de d?tails sur ce processus, voir [Enregistrer un compl?ment Office qui utilise l'authentification unique aupr?s du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="4d551-157">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="4d551-158">Configurer le complément</span><span class="sxs-lookup"><span data-stu-id="4d551-158">Configure the add-in</span></span>

<span data-ttu-id="4d551-159">Ajoutez un nouveau balisage au manifeste du complément :</span><span class="sxs-lookup"><span data-stu-id="4d551-159">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="4d551-160">**WebApplicationInfo**: le parent des éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="4d551-160">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="4d551-161">**Id** - ID du client du compl?ment : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="4d551-161">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="4d551-162">Voir[Enregistrer un complément Office utilisant une SSO (authentification unique) avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="4d551-162">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="4d551-163">**Ressource**: l’URL du complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-163">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="4d551-164">Il s’agit du même URI (y compris le protocole`api:`) que vous avez utilisé lors de l’inscription du complément dans AAD.</span><span class="sxs-lookup"><span data-stu-id="4d551-164">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="4d551-165">Le domaine et les sous-domaines doivent être les mêmes que ceux utilisés dans les URLs dans la section`<Resources>` du manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-165">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="4d551-166">**Scopes**: le parent d’un ou plusieurs éléments **Scope**.</span><span class="sxs-lookup"><span data-stu-id="4d551-166">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="4d551-167">**Scope**: spécifie une autorisation nécessaire pour le complément dans l’AAD.</span><span class="sxs-lookup"><span data-stu-id="4d551-167">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="4d551-168">L' `profile` autorisation est toujours n?cessaire et il peut s'agir de la seule autorisation n?cessaire si votre compl?ment n'acc?de pas ? Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="4d551-168">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="4d551-169">Si c'est le cas, vous avez ?galement besoin des ?l?ments d'une **?tendue**pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="4d551-169">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="4d551-170">Les biblioth?ques que vous utilisez dans votre code pour acc?der ? Microsoft Graph peuvent avoir des besoin d'autorisations suppl?mentaires.</span><span class="sxs-lookup"><span data-stu-id="4d551-170">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="4d551-171">Par exemple, Microsoft Authentication Library (MSAL) pour .NET n?cessite `offline_access` une autorisation.</span><span class="sxs-lookup"><span data-stu-id="4d551-171">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="4d551-172">Pour plus d'informations, voir [Autoriser Microsoft Graph ? partir d'un compl?ment Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="4d551-172">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="4d551-p115">Pour les hôtes Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="4d551-p115">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="4d551-175">Voici un exemple de marques de révision :</span><span class="sxs-lookup"><span data-stu-id="4d551-175">The following is an example of the markup:</span></span>

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a><span data-ttu-id="4d551-176">Ajouter du code côté client</span><span class="sxs-lookup"><span data-stu-id="4d551-176">Add client-side code</span></span>

<span data-ttu-id="4d551-177">Ajoutez un code JavaScript pour le complément à :</span><span class="sxs-lookup"><span data-stu-id="4d551-177">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="4d551-178">Appelez[getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="4d551-178">Call [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="4d551-179">Analyser le jeton d’accès ou le transmettre au code côté serveur du complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-179">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="4d551-180">Voici un exemple simple d’un appel à`getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="4d551-180">Here's a simple example of a call to `getAccessTokenAsync`.</span></span>

> [!NOTE]
> <span data-ttu-id="4d551-181">Cet exemple ne pr?sente explicitement qu'un seul type d'erreur.</span><span class="sxs-lookup"><span data-stu-id="4d551-181">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="4d551-182">Pour avoir des exemples de traitement des erreurs plus ?labor?s, voir [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) et [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="4d551-182">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="4d551-183">Et voir [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="4d551-183">And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

<span data-ttu-id="4d551-184">Voici un exemple simple d?un passage de token du compl?ment vers le serveur.</span><span class="sxs-lookup"><span data-stu-id="4d551-184">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="4d551-185">Le token est inclus en tant qu' `Authorization` en-t?te lors de l'envoi d'une demande au serveur.</span><span class="sxs-lookup"><span data-stu-id="4d551-185">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="4d551-186">Dans cet exemple, l'envoi de donn?es JSON se fait en utilisant la m?thode `POST`, mais `GET` est suffisant pour envoyer le token d'acc?s lorsque vous n'?crivez pas sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="4d551-186">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a><span data-ttu-id="4d551-187">Quand appeler la méthode</span><span class="sxs-lookup"><span data-stu-id="4d551-187">When to call the method</span></span>

<span data-ttu-id="4d551-188">Si votre complément ne peut pas être utilisé lorsqu’ aucun utilisateur n’est connecté à Office, vous devez alors appeler`getAccessTokenAsync` \* au lancement du complément\*.</span><span class="sxs-lookup"><span data-stu-id="4d551-188">If your add-in cannot be used when no user is logged into Office, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="4d551-189">Si le complément possède certaines fonctionnalités qui ne nécessitent pas un accès à l’utilisateur, ensuite appelez`getAccessTokenAsync`\* lorsque l’utilisateur effectue une action qui requiert un utilisateur connecté\*.</span><span class="sxs-lookup"><span data-stu-id="4d551-189">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="4d551-190">Les appels répétés à `getAccessTokenAsync` ne causent aucune dégradation importante des performances, car Office met en cache le jeton d’accès et le réutilise jusqu'à ce qu’il arrive à expiration, sans effectuer un autre appel à l’AAD V.</span><span class="sxs-lookup"><span data-stu-id="4d551-190">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="4d551-191">Point de terminaison 2.0 dès que `getAccessTokenAsync` est appelé.</span><span class="sxs-lookup"><span data-stu-id="4d551-191">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="4d551-192">Ainsi, vous pouvez ajouter des appels de `getAccessTokenAsync` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le jeton est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="4d551-192">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="4d551-193">Ajouter du code côté serveur</span><span class="sxs-lookup"><span data-stu-id="4d551-193">Add server-side code</span></span>

<span data-ttu-id="4d551-194">Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement.</span><span class="sxs-lookup"><span data-stu-id="4d551-194">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="4d551-195">Quelques t?ches c?t? serveur que votre compl?ment pourrait faire :</span><span class="sxs-lookup"><span data-stu-id="4d551-195">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="4d551-196">Cr?er d'une ou plusieurs m?thodes d'API Web qui utilisent des informations sur l'utilisateur qui sont extraitent du token ; par exemple, une m?thode qui recherche les pr?f?rences de l'utilisateur dans votre base de donn?es h?berg?e.</span><span class="sxs-lookup"><span data-stu-id="4d551-196">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="4d551-197">(Voir **Utilisation du token SSO en tant qu'identit?** ci-dessous). En fonction de votre langue et de votre structure, des biblioth?ques peuvent ?tre disponibles pour simplifier le code que vous devez ?crire.</span><span class="sxs-lookup"><span data-stu-id="4d551-197">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="4d551-198">Obtenir des donn?es Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="4d551-198">Get Microsoft Graph data.</span></span> <span data-ttu-id="4d551-199">Votre code côté serveur doit effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="4d551-199">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="4d551-200">Valider le jeton d’accès (voir**Valider du jeton d’accès** ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="4d551-200">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="4d551-201">Démarrer le flux « de la part de » avec un appel du point de terminaison Azure AD v2.0 qui inclut le jeton d’accès du complément, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et code secret).</span><span class="sxs-lookup"><span data-stu-id="4d551-201">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="4d551-202">Dans ce contexte, le jeton d’accès est appelé le jeton bootstrap.</span><span class="sxs-lookup"><span data-stu-id="4d551-202">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="4d551-203">Mettre en cache le nouveau token renvoy? par le flux ? de la part de ?.</span><span class="sxs-lookup"><span data-stu-id="4d551-203">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="4d551-204">Obtenir des données à partir de Microsoft Graph en utilisant le nouveau jeton.</span><span class="sxs-lookup"><span data-stu-id="4d551-204">Get data from Microsoft Graph by using the new token.</span></span>

 <span data-ttu-id="4d551-205">Pour plus de d?tails sur l'obtention d'un acc?s autoris? aux donn?es Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre compl?ment Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="4d551-205">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="4d551-206">Valider le jeton d’accès</span><span class="sxs-lookup"><span data-stu-id="4d551-206">Validate the access token</span></span>

<span data-ttu-id="4d551-207">Quand l’API web reçoit le jeton d’accès, elle doit valider son fonctionnement avant de l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="4d551-207">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="4d551-208">Le jeton est un jeton JWT. En d’autres termes, la validation se déroule comme dans la plupart des flux OAuth standard.</span><span class="sxs-lookup"><span data-stu-id="4d551-208">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="4d551-209">Il existe un certain nombre de bibliothèques pouvant gérer la validation JWT qui sont toutes, au minimum, chargées de :</span><span class="sxs-lookup"><span data-stu-id="4d551-209">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="4d551-210">vérifier que le jeton est bien formé ;</span><span class="sxs-lookup"><span data-stu-id="4d551-210">Checking that the token is well-formed</span></span>
- <span data-ttu-id="4d551-211">vérifier que le jeton a été émis par l’autorité souhaitée ;</span><span class="sxs-lookup"><span data-stu-id="4d551-211">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="4d551-212">vérifier que le jeton est destiné à l’API web.</span><span class="sxs-lookup"><span data-stu-id="4d551-212">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="4d551-213">Suivez les recommandations suivantes quand vous validez le jeton :</span><span class="sxs-lookup"><span data-stu-id="4d551-213">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="4d551-214">Les jetons SSO valides doivent être émis par l’autorité Azure `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="4d551-214">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="4d551-215">La revendication `iss` dans le jeton doit commencer par cette valeur.</span><span class="sxs-lookup"><span data-stu-id="4d551-215">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="4d551-216">Le paramètre `aud` du jeton devra correspondre à l’ID d’application de l’enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-216">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="4d551-217">Le paramètre `scp` du jeton devra correspondre à `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="4d551-217">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="4d551-218">Utilisation du jeton SSO comme identité</span><span class="sxs-lookup"><span data-stu-id="4d551-218">Using the SSO token as an identity</span></span>

<span data-ttu-id="4d551-219">Si votre complément doit vérifier l’identité de l’utilisateur, le jeton SSO contient des informations utiles pour établir son identité.</span><span class="sxs-lookup"><span data-stu-id="4d551-219">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="4d551-220">Les revendications suivantes présentes dans le jeton concernent l’identité de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d551-220">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="4d551-221">`name`: le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d551-221">`name` - The user's display name.</span></span>
- <span data-ttu-id="4d551-222">`preferred_username`: l’adresse de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4d551-222">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="4d551-223">`oid` : un GUID représentant l’ID de l’utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="4d551-223">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="4d551-224">`tid`: un GUID représentant l’ID de l’organisation de l’utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="4d551-224">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="4d551-225">Étant donné que les valeurs `name` et `preferred_username` peuvent être amenées à changer, nous vous recommandons d’utiliser les valeurs `oid` et `tid` pour corréler l’identité de l’utilisateur avec le service d’autorisation de votre API principale.</span><span class="sxs-lookup"><span data-stu-id="4d551-225">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="4d551-226">Par exemple, votre service peut mettre en forme ces valeurs de la façon suivante `{oid-value}@{tid-value}`, puis stocker cette mise en forme sous forme de valeur dans l’enregistrement de l’utilisateur dans votre base de données utilisateur interne.</span><span class="sxs-lookup"><span data-stu-id="4d551-226">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="4d551-227">Lors des demandes ultérieures, l’utilisateur pourra être récupéré grâce à cette valeur et l’accès à certaines ressources pourra être déterminé selon les mécanismes de contrôle d’accès existants.</span><span class="sxs-lookup"><span data-stu-id="4d551-227">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="4d551-228">Exemple de token</span><span class="sxs-lookup"><span data-stu-id="4d551-228">Example access token</span></span>

<span data-ttu-id="4d551-229">Voici une charge utile d?cod?e typique de token.</span><span class="sxs-lookup"><span data-stu-id="4d551-229">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="4d551-230">Pour plus d’informations sur les propriétés, voir[jetons référence (token) version 2.0 Azure Active Directory](/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="4d551-230">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",
    iat: 1521143967,
    nbf: 1521143967,
    exp: 1521147867,
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",
    azpacr: "0",
    e_exp: 262800,
    name: "Mila Nikolova",
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",
    preferred_username: "milan@contoso.com",
    scp: "access_as_user",
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",
    uti: "MICAQyhrH02ov54bCtIDAA",
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="4d551-231">Utilisation de l’authentification unique SSO en accompagnement d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="4d551-231">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="4d551-232">Il existe quelques différences mineures, mais importantes, en ce qui concerne l'utilisation de la connexion unique SSO dans un complément Outlook à partir de son utilisation dans un complément Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="4d551-232">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="4d551-233">Assurez-vous de lire [Authentifier un utilisateur avec un token unique log? dans le compl?ment Outlook](/outlook/add-ins/authenticate-a-user-with-an-sso-token) et l' [?tude de cas : Impl?menter la connexion unique ? votre service dans un compl?ment Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="4d551-233">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="4d551-234">RÉFÉRENCE D’API D’AUTHENTIFICATION UNIQUE SSO</span><span class="sxs-lookup"><span data-stu-id="4d551-234">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="4d551-235">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4d551-235">getAccessTokenAsync</span></span>

<span data-ttu-id="4d551-236">L’espace de noms d’authentification Office,`Office.context.auth`, fournit une méthode,`getAccessTokenAsync` qui permet à l’hôte Office d’obtenir le jeton d’accès à l’application web du complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-236">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="4d551-237">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="4d551-237">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="4d551-238">Cette méthode appelle le point de terminaison Azure Active Directory V 2.0 pour obtenir un jeton d’accès à l’application web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="4d551-238">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="4d551-239">Ceci permet à des compléments d’identifier les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="4d551-239">This enables add-ins to identify users.</span></span> <span data-ttu-id="4d551-240">Le Code côté serveur peut utiliser ce jeton pour accéder à Microsoft Graph pour l’application web du complément à l’aide du [flux OAuth « Pour le compte de »](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="4d551-240">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="4d551-241">Dans Outlook, cette API n'est pas prise en charge si le complément est chargé dans une boîte aux lettres Outlook.com ou Gmail.</span><span class="sxs-lookup"><span data-stu-id="4d551-241">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="4d551-242">Hôtes</span><span class="sxs-lookup"><span data-stu-id="4d551-242">Hosts</span></span>|<span data-ttu-id="4d551-243">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="4d551-243">Excel, OneNote, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="4d551-244">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="4d551-244">Requirement sets</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)|[<span data-ttu-id="4d551-245">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="4d551-245">IdentityAPI</span></span>](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)|

#### <a name="parameters"></a><span data-ttu-id="4d551-246">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4d551-246">Parameters</span></span>

<span data-ttu-id="4d551-247">`options`: Facultatif.</span><span class="sxs-lookup"><span data-stu-id="4d551-247">`options` - Optional.</span></span> <span data-ttu-id="4d551-248">Accepte un objet `AuthOptions` (voir ci-dessous) pour définir les comportements d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="4d551-248">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="4d551-249">`callback`: Facultatif.</span><span class="sxs-lookup"><span data-stu-id="4d551-249">`callback` - Optional.</span></span> <span data-ttu-id="4d551-250">Accepte une méthode de rappel qui peut analyser le jeton pour l’ID de l’utilisateur ou utilisez le jeton dans le flux de « de la part de » pour accéder à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="4d551-250">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="4d551-251">Si[AsyncResult](/javascript/api/office/office.asyncresult) `.status` est « Réussi », puis`AsyncResult.value` est le AAD v brut.</span><span class="sxs-lookup"><span data-stu-id="4d551-251">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="4d551-252">2.0: le jeton d’accès mis en forme.</span><span class="sxs-lookup"><span data-stu-id="4d551-252">2.0-formatted access token.</span></span>

<span data-ttu-id="4d551-253">L’interface`AuthOptions` fournit des options pour l’expérience utilisateur lorsqu’Office obtient un jeton d’accès pour le complément à partir d’AAD v.</span><span class="sxs-lookup"><span data-stu-id="4d551-253">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="4d551-254">2.0 avec la méthode`getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="4d551-254">2.0 with the `getAccessTokenAsync` method.</span></span>

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try"
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```
