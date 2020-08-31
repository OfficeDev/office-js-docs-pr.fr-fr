---
title: Activer l’authentification unique pour des compléments Office
description: Découvrez comment activer l’authentification unique pour les Compléments Office à l’aide de votre compte courant Microsoft personnel, professionnel ou scolaire.
ms.date: 07/30/2020
localization_priority: Priority
ms.openlocfilehash: ec4fc9f91f3cbc9f8882ed491c7c5bc68be346ed
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293197"
---
# <a name="enable-single-sign-on-for-office-add-ins"></a><span data-ttu-id="2437d-103">Activer la connexion unique pour des compléments Office</span><span class="sxs-lookup"><span data-stu-id="2437d-103">Enable single sign-on for Office Add-ins</span></span>


<span data-ttu-id="2437d-p101">Les utilisateurs se connectent à Office (plateformes en ligne, mobiles ou de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou Microsoft 365 Éducation. Vous pouvez en tirer parti et utiliser l’authentification unique (SSO) pour autoriser l’utilisateur à accéder à votre complément sans qu’il doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="2437d-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their Microsoft 365 Education or work account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Image illustrant le processus de connexion pour un complément](../images/sso-for-office-addins.png)

## <a name="requirements-and-best-practices"></a><span data-ttu-id="2437d-107">Meilleures Pratiques et Conditions Requises</span><span class="sxs-lookup"><span data-stu-id="2437d-107">Requirements and Best Practices</span></span>

<span data-ttu-id="2437d-108">Si vous travaillez avec un complément **Outlook**, assurez-vous d'activer l'authentification moderne pour la location de Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="2437d-108">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="2437d-109">Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="2437d-109">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="2437d-110">Vous ne devez*pas* dépendre de l’authentification unique SSO comme seule méthode de votre complément d’authentification.</span><span class="sxs-lookup"><span data-stu-id="2437d-110">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="2437d-111">Vous devez implémenter un système d’authentification secondaire vers lequel votre complément peut revenir dans certaines situations d’erreur.</span><span class="sxs-lookup"><span data-stu-id="2437d-111">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="2437d-112">Vous pouvez utiliser un système de tableaux d’utilisateur et d’authentification, ou vous pouvez tirer parti d’un des fournisseurs de connexion sociale.</span><span class="sxs-lookup"><span data-stu-id="2437d-112">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="2437d-113">Pour plus d’informations sur la procédure à suivre avec un complément Office, voir[Services externes autorisées dans votre complément Office](auth-external-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-113">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).</span></span> <span data-ttu-id="2437d-114">Concernant*Outlook*, il existe un système de secours recommandé.</span><span class="sxs-lookup"><span data-stu-id="2437d-114">For *Outlook*, there is a recommended fallback system.</span></span> <span data-ttu-id="2437d-115">Pour plus d’informations, voir[Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook](../outlook/implement-sso-in-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-115">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md).</span></span> <span data-ttu-id="2437d-116">Pour consulter des exemples d’utilisation d’Azure Active Directory comme système de secours, voir [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) et [SSO ASP.NET pour complément Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="2437d-116">For samples that use Azure Active Directory as the fallback system, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>

## <a name="how-sso-works-at-runtime"></a><span data-ttu-id="2437d-117">Mode de fonctionnement de l’authentification unique SSO en cours d’exécution</span><span class="sxs-lookup"><span data-stu-id="2437d-117">How SSO works at runtime</span></span>

<span data-ttu-id="2437d-118">Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique SSO.</span><span class="sxs-lookup"><span data-stu-id="2437d-118">The following diagram shows how the SSO process works.</span></span>

![Un diagramme illustrant le processus d’authentification unique](../images/sso-overview-diagram.png)

1. <span data-ttu-id="2437d-120">Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="2437d-120">In the add-in, JavaScript calls a new Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span> <span data-ttu-id="2437d-121">Cela indique à l’application cliente Office qu’elle doit obtenir un jeton d’accès au complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-121">This tells the Office client application to obtain an access token to the add-in.</span></span> <span data-ttu-id="2437d-122">Voir [Exemple de token](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="2437d-122">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="2437d-123">Si l’utilisateur n’est pas connecté, l’application cliente Office ouvre une fenêtre contextuelle pour qu’il se connecte.</span><span class="sxs-lookup"><span data-stu-id="2437d-123">If the user is not signed in, the Office client application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="2437d-124">Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.</span><span class="sxs-lookup"><span data-stu-id="2437d-124">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="2437d-125">L’application cliente Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="2437d-125">The Office client application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="2437d-126">Azure AD envoie le jeton de complément à l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="2437d-126">Azure AD sends the add-in token to the Office client application.</span></span>
6. <span data-ttu-id="2437d-127">L’application cliente Office envoie le**jeton de complément (token)** au complément dans le cadre de l’objet de résultat renvoyé par l’appel`getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="2437d-127">The Office client application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessToken` call.</span></span>
7. <span data-ttu-id="2437d-128">Dans le compl?ment, JavaScript peut analyser le token et extraire les informations dont il a besoin, telles que l'adresse e-mail de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2437d-128">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span>
8. <span data-ttu-id="2437d-129">Optionnellement, le compl?ment peut envoyer une requ?te HTTP ? son serveur pour obtenir plus de donn?es sur l'utilisateur, notamment les pr?f?rences de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2437d-129">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="2437d-130">Alternativement, le token lui-m?me pourrait ?tre envoy? au serveur pour analyse et validation.</span><span class="sxs-lookup"><span data-stu-id="2437d-130">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="2437d-131">Développer un complément d’authentification unique SSO</span><span class="sxs-lookup"><span data-stu-id="2437d-131">Develop an SSO add-in</span></span>

<span data-ttu-id="2437d-132">Cette section décrit les tâches impliquées dans la création d’un complément Office qui utilise l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="2437d-132">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="2437d-133">Ces tâches sont décrites ici indépendamment du langage et de l’infrastructure.</span><span class="sxs-lookup"><span data-stu-id="2437d-133">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="2437d-134">Pour consulter des procédures pas à pas détaillées, voir :</span><span class="sxs-lookup"><span data-stu-id="2437d-134">For detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="2437d-135">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="2437d-135">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="2437d-136">Créer un complément Office ASP.NET qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="2437d-136">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

> [!NOTE]
> <span data-ttu-id="2437d-137">Vous pouvez utiliser le générateur Yeoman pour créer votre complément Office compatible avec l’authentification unique, Node.js..</span><span class="sxs-lookup"><span data-stu-id="2437d-137">You can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="2437d-138">Le générateur Yeoman simplifie le processus de création d’un complément avec authentification unique en automatisant les étapes nécessaires pour configurer l’authentification unique dans Azure et la génération du code nécessaire pour qu’un complément utilise l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="2437d-138">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="2437d-139">Pour plus d'informations, consultez [Démarrage rapide de l'authentification unique](../quickstarts/sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-139">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

### <a name="create-the-service-application"></a><span data-ttu-id="2437d-140">Créer l’application de service</span><span class="sxs-lookup"><span data-stu-id="2437d-140">Create the service application</span></span>

<span data-ttu-id="2437d-p108">Enregistrer le complément auprès du portail d’inscription pour le point de terminaison Azure v2.0. Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :</span><span class="sxs-lookup"><span data-stu-id="2437d-p108">Register the add-in at the registration portal for the Azure v2.0 endpoint. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="2437d-143">Obtenez un ID client et un code secret pour le complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-143">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="2437d-144">Spécifiez les autorisations dont votre complément a besoin pour AAD v.</span><span class="sxs-lookup"><span data-stu-id="2437d-144">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="2437d-145"> Point de terminaison 2.0 (et ?ventuellement Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="2437d-145">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="2437d-146">L'autorisation "profil" est toujours n?cessaire.</span><span class="sxs-lookup"><span data-stu-id="2437d-146">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="2437d-147">Accordez la confiance de l’application cliente Office au complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-147">Grant the Office client application trust to the add-in.</span></span>
* <span data-ttu-id="2437d-148">Pré-autorisez l’application cliente Office pour le complément avec l’autorisation par défaut *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="2437d-148">Preauthorize the Office client application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="2437d-149">Pour plus de d?tails sur ce processus, voir [Enregistrer un compl?ment Office qui utilise l'authentification unique aupr?s du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-149">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="2437d-150">Configurer le complément</span><span class="sxs-lookup"><span data-stu-id="2437d-150">Configure the add-in</span></span>

<span data-ttu-id="2437d-151">Ajoutez un nouveau balisage au manifeste du complément :</span><span class="sxs-lookup"><span data-stu-id="2437d-151">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="2437d-152">**WebApplicationInfo**: le parent des éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="2437d-152">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="2437d-153">**Id** - ID du client du compl?ment : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="2437d-153">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="2437d-154">Voir[Enregistrer un complément Office utilisant une SSO (authentification unique) avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-154">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="2437d-155">**Ressource**: l’URL du complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-155">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="2437d-156">Il s’agit du même URI (y compris le protocole`api:`) que vous avez utilisé lors de l’inscription du complément dans AAD.</span><span class="sxs-lookup"><span data-stu-id="2437d-156">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="2437d-157">Le domaine et les sous-domaines doivent être les mêmes que ceux utilisés dans les URLs dans la section`<Resources>` du manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-157">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="2437d-158">**Scopes**: le parent d’un ou plusieurs éléments **Scope**.</span><span class="sxs-lookup"><span data-stu-id="2437d-158">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="2437d-159">**Scope**: spécifie une autorisation nécessaire pour le complément dans l’AAD.</span><span class="sxs-lookup"><span data-stu-id="2437d-159">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="2437d-160">L' `profile` autorisation est toujours n?cessaire et il peut s'agir de la seule autorisation n?cessaire si votre compl?ment n'acc?de pas ? Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2437d-160">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="2437d-161">Si c'est le cas, vous avez ?galement besoin des ?l?ments d'une **?tendue**pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="2437d-161">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="2437d-162">Les biblioth?ques que vous utilisez dans votre code pour acc?der ? Microsoft Graph peuvent avoir des besoin d'autorisations suppl?mentaires.</span><span class="sxs-lookup"><span data-stu-id="2437d-162">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="2437d-163">Par exemple, Microsoft Authentication Library (MSAL) pour .NET n?cessite `offline_access` une autorisation.</span><span class="sxs-lookup"><span data-stu-id="2437d-163">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="2437d-164">Pour plus d'informations, voir [Autoriser Microsoft Graph ? partir d'un compl?ment Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-164">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="2437d-p113">Pour les applications Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="2437d-p113">For Office applications other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="2437d-167">Voici un exemple de marques de révision :</span><span class="sxs-lookup"><span data-stu-id="2437d-167">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="2437d-168">Ajouter du code côté client</span><span class="sxs-lookup"><span data-stu-id="2437d-168">Add client-side code</span></span>

<span data-ttu-id="2437d-169">Ajoutez un code JavaScript pour le complément à :</span><span class="sxs-lookup"><span data-stu-id="2437d-169">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="2437d-170">Appelez [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="2437d-170">Call [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span>

* <span data-ttu-id="2437d-171">Analyser le jeton d’accès ou le transmettre au code côté serveur du complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-171">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="2437d-172">Voici un exemple simple d’un appel à`getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="2437d-172">Here's a simple example of a call to `getAccessToken`.</span></span>

> [!NOTE]
> <span data-ttu-id="2437d-173">Cet exemple ne pr?sente explicitement qu'un seul type d'erreur.</span><span class="sxs-lookup"><span data-stu-id="2437d-173">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="2437d-174">Pour des exemples de traitement des erreurs plus élaborés, voir [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) et [SSO ASP.NET pour complément Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="2437d-174">For examples of more elaborate error handling, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>


```js
async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken();

        // The /api/DoSomething controller will make the token exchange and use the
        // access token it gets back to make the call to MS Graph.
        getData("/api/DoSomething", bootstrapToken);
    }
    catch (exception) {
        if (exception.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // Microsoft 365 Education or work account, or a Microsoft account.
        } else {
            // Handle error
        }
    }
}
```

<span data-ttu-id="2437d-175">Voici un exemple simple d?un passage de token du compl?ment vers le serveur.</span><span class="sxs-lookup"><span data-stu-id="2437d-175">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="2437d-176">Le token est inclus en tant qu' `Authorization` en-t?te lors de l'envoi d'une demande au serveur.</span><span class="sxs-lookup"><span data-stu-id="2437d-176">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="2437d-177">Dans cet exemple, l'envoi de donn?es JSON se fait en utilisant la m?thode `POST`, mais `GET` est suffisant pour envoyer le token d'acc?s lorsque vous n'?crivez pas sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="2437d-177">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + bootstrapToken
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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="2437d-178">Quand appeler la méthode</span><span class="sxs-lookup"><span data-stu-id="2437d-178">When to call the method</span></span>

<span data-ttu-id="2437d-179">Si votre complément ne peut pas être utilisé lorsqu’aucun utilisateur n’est connecté à Office, vous devez alors appeler`getAccessToken` \* au lancement du complément\* et passer `allowSignInPrompt: true` dans le `options` paramètre de `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="2437d-179">If your add-in cannot be used when there is no user currently logged into Office, then you should call `getAccessToken` *when the add-in launches* and pass `allowSignInPrompt: true` in the `options` parameter of `getAccessToken`.</span></span> <span data-ttu-id="2437d-180">Par exemple: `OfficeRuntime.auth.getAccessToken( { allowSignInPrompt: true });`</span><span class="sxs-lookup"><span data-stu-id="2437d-180">For example; `OfficeRuntime.auth.getAccessToken( { allowSignInPrompt: true });`</span></span>

<span data-ttu-id="2437d-181">Si le complément possède certaines fonctionnalités qui ne nécessitent pas un accès à l’utilisateur, ensuite appelez`getAccessToken`\* lorsque l’utilisateur effectue une action qui requiert un utilisateur connecté\*.</span><span class="sxs-lookup"><span data-stu-id="2437d-181">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessToken` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="2437d-182">Les appels répétés à `getAccessToken` ne causent aucune dégradation importante des performances, car Office met en cache le jeton d’amorçage et le réutilise jusqu'à ce qu’il arrive à expiration, sans effectuer un autre appel vers l’AAD v.</span><span class="sxs-lookup"><span data-stu-id="2437d-182">There is no significant performance degradation with redundant calls of `getAccessToken` because Office caches the bootstrap token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="2437d-183">Point de terminaison 2.0 dès que `getAccessToken` est appelé.</span><span class="sxs-lookup"><span data-stu-id="2437d-183">2.0 endpoint whenever `getAccessToken` is called.</span></span> <span data-ttu-id="2437d-184">Ainsi, vous pouvez ajouter des appels de `getAccessToken` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le jeton est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="2437d-184">So you can add calls of `getAccessToken` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="2437d-185">Ajouter du code côté serveur</span><span class="sxs-lookup"><span data-stu-id="2437d-185">Add server-side code</span></span>

<span data-ttu-id="2437d-186">Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement.</span><span class="sxs-lookup"><span data-stu-id="2437d-186">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="2437d-187">Quelques t?ches c?t? serveur que votre compl?ment pourrait faire :</span><span class="sxs-lookup"><span data-stu-id="2437d-187">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="2437d-188">Cr?er d'une ou plusieurs m?thodes d'API Web qui utilisent des informations sur l'utilisateur qui sont extraitent du token ; par exemple, une m?thode qui recherche les pr?f?rences de l'utilisateur dans votre base de donn?es h?berg?e.</span><span class="sxs-lookup"><span data-stu-id="2437d-188">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="2437d-189">(Voir **Utilisation du token SSO en tant qu'identit?** ci-dessous). En fonction de votre langue et de votre structure, des biblioth?ques peuvent ?tre disponibles pour simplifier le code que vous devez ?crire.</span><span class="sxs-lookup"><span data-stu-id="2437d-189">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="2437d-190">Obtenir des donn?es Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2437d-190">Get Microsoft Graph data.</span></span> <span data-ttu-id="2437d-191">Votre code côté serveur doit effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="2437d-191">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="2437d-192">Démarrer le flux « de la part de » avec un appel du point de terminaison Azure AD v2.0 qui inclut le jeton d’accès du complément, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et code secret).</span><span class="sxs-lookup"><span data-stu-id="2437d-192">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="2437d-193">Dans ce contexte, le jeton d’accès est appelé le jeton bootstrap.</span><span class="sxs-lookup"><span data-stu-id="2437d-193">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="2437d-194">Obtenir des données à partir de Microsoft Graph en utilisant le nouveau jeton.</span><span class="sxs-lookup"><span data-stu-id="2437d-194">Get data from Microsoft Graph by using the new token.</span></span>
    * <span data-ttu-id="2437d-195">Si vous le souhaitez, avant de lancer le flux, validez le jeton d’accès (voir **Valider le jeton d’accès** ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="2437d-195">Optionally, before initiating the flow, validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="2437d-196">Si vous le souhaitez, une fois l’exécution du flux on-behalf-of terminée, mettez en cache le nouveau jeton d’accès renvoyé à partir du flux de façon à ce qu’il soit réutilisé dans d’autres appels à Microsoft Graph jusqu’à son expiration.</span><span class="sxs-lookup"><span data-stu-id="2437d-196">Optionally, after the on-behalf-of flow completes, cache the new access token that is returned from the flow so that it an be reused in other calls to Microsoft Graph until it expires.</span></span>

 <span data-ttu-id="2437d-197">Pour plus de d?tails sur l'obtention d'un acc?s autoris? aux donn?es Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre compl?ment Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-197">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="2437d-198">Valider le jeton d’accès</span><span class="sxs-lookup"><span data-stu-id="2437d-198">Validate the access token</span></span>

<span data-ttu-id="2437d-199">Quand l’API web reçoit le jeton d’accès, elle peut valider son fonctionnement avant de l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="2437d-199">Once the Web API receives the access token, it can validate it before using it.</span></span> <span data-ttu-id="2437d-200">Le jeton est un jeton JWT. En d’autres termes, la validation se déroule comme dans la plupart des flux OAuth standard.</span><span class="sxs-lookup"><span data-stu-id="2437d-200">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="2437d-201">Il existe un certain nombre de bibliothèques pouvant gérer la validation JWT qui sont toutes, au minimum, chargées de :</span><span class="sxs-lookup"><span data-stu-id="2437d-201">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="2437d-202">vérifier que le jeton est bien formé ;</span><span class="sxs-lookup"><span data-stu-id="2437d-202">Checking that the token is well-formed</span></span>
- <span data-ttu-id="2437d-203">vérifier que le jeton a été émis par l’autorité souhaitée ;</span><span class="sxs-lookup"><span data-stu-id="2437d-203">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="2437d-204">vérifier que le jeton est destiné à l’API web.</span><span class="sxs-lookup"><span data-stu-id="2437d-204">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="2437d-205">Suivez les recommandations suivantes quand vous validez le jeton :</span><span class="sxs-lookup"><span data-stu-id="2437d-205">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="2437d-206">Les jetons SSO valides doivent être émis par l’autorité Azure `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="2437d-206">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="2437d-207">La revendication `iss` dans le jeton doit commencer par cette valeur.</span><span class="sxs-lookup"><span data-stu-id="2437d-207">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="2437d-208">Le paramètre `aud` du jeton devra correspondre à l’ID d’application de l’enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-208">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="2437d-209">Le paramètre `scp` du jeton devra correspondre à `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="2437d-209">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="2437d-210">Utilisation du jeton SSO comme identité</span><span class="sxs-lookup"><span data-stu-id="2437d-210">Using the SSO token as an identity</span></span>

<span data-ttu-id="2437d-211">Si votre complément doit vérifier l’identité de l’utilisateur, le jeton SSO contient des informations utiles pour établir son identité.</span><span class="sxs-lookup"><span data-stu-id="2437d-211">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="2437d-212">Les revendications suivantes présentes dans le jeton concernent l’identité de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2437d-212">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="2437d-213">`name`: le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2437d-213">`name` - The user's display name.</span></span>
- <span data-ttu-id="2437d-214">`preferred_username`: l’adresse de messagerie de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2437d-214">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="2437d-215">`oid` : un GUID représentant l’ID de l’utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="2437d-215">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="2437d-216">`tid`: un GUID représentant l’ID de l’organisation de l’utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="2437d-216">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="2437d-217">Étant donné que les valeurs `name` et `preferred_username` peuvent être amenées à changer, nous vous recommandons d’utiliser les valeurs `oid` et `tid` pour corréler l’identité de l’utilisateur avec le service d’autorisation de votre API principale.</span><span class="sxs-lookup"><span data-stu-id="2437d-217">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="2437d-218">Par exemple, votre service peut mettre en forme ces valeurs de la façon suivante `{oid-value}@{tid-value}`, puis stocker cette mise en forme sous forme de valeur dans l’enregistrement de l’utilisateur dans votre base de données utilisateur interne.</span><span class="sxs-lookup"><span data-stu-id="2437d-218">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="2437d-219">Lors des demandes ultérieures, l’utilisateur pourra être récupéré grâce à cette valeur et l’accès à certaines ressources pourra être déterminé selon les mécanismes de contrôle d’accès existants.</span><span class="sxs-lookup"><span data-stu-id="2437d-219">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="2437d-220">Exemple de token</span><span class="sxs-lookup"><span data-stu-id="2437d-220">Example access token</span></span>

<span data-ttu-id="2437d-221">Voici une charge utile d?cod?e typique de token.</span><span class="sxs-lookup"><span data-stu-id="2437d-221">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="2437d-222">Pour plus d’informations sur les propriétés, voir[jetons référence (token) version 2.0 Azure Active Directory](/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="2437d-222">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>

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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="2437d-223">Utilisation de l’authentification unique SSO en accompagnement d’un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="2437d-223">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="2437d-224">Il existe quelques différences mineures, mais importantes, en ce qui concerne l'utilisation de la connexion unique SSO dans un complément Outlook à partir de son utilisation dans un complément Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="2437d-224">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="2437d-225">Assurez-vous de lire [Authentifier un utilisateur avec un token unique log? dans le compl?ment Outlook](../outlook/authenticate-a-user-with-an-sso-token.md) et l' [?tude de cas : Impl?menter la connexion unique ? votre service dans un compl?ment Outlook](../outlook/implement-sso-in-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="2437d-225">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](../outlook/authenticate-a-user-with-an-sso-token.md) and [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="2437d-226">RÉFÉRENCE D’API D’AUTHENTIFICATION UNIQUE SSO</span><span class="sxs-lookup"><span data-stu-id="2437d-226">SSO API reference</span></span>

### <a name="getaccesstoken"></a><span data-ttu-id="2437d-227">getAccessToken</span><span class="sxs-lookup"><span data-stu-id="2437d-227">getAccessToken</span></span>

<span data-ttu-id="2437d-228">L'espace de noms OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth), `OfficeRuntime.Auth`, fournit une méthode `getAccessToken` qui permet à l'application Office d'obtenir un jeton d'accès à l'application web du module complémentaire.</span><span class="sxs-lookup"><span data-stu-id="2437d-228">The OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth) namespace, `OfficeRuntime.Auth`, provides a method, `getAccessToken` that enables the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="2437d-229">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="2437d-229">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessToken(options?: AuthOptions: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="2437d-230">Cette méthode appelle le point de terminaison Azure Active Directory V 2.0 pour obtenir un jeton d’accès à l’application web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="2437d-230">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="2437d-231">Ceci permet à des compléments d’identifier les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="2437d-231">This enables add-ins to identify users.</span></span> <span data-ttu-id="2437d-232">Le Code côté serveur peut utiliser ce jeton pour accéder à Microsoft Graph pour l’application web du complément à l’aide du [flux OAuth « Pour le compte de »](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="2437d-232">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="2437d-233">Dans Outlook, cette API n'est pas prise en charge si le complément est chargé dans une boîte aux lettres Outlook.com ou Gmail.</span><span class="sxs-lookup"><span data-stu-id="2437d-233">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="2437d-234">Hôtes</span><span class="sxs-lookup"><span data-stu-id="2437d-234">Hosts</span></span>|<span data-ttu-id="2437d-235">Excel, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="2437d-235">Excel, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="2437d-236">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2437d-236">Requirement sets</span></span>](specify-office-hosts-and-api-requirements.md)|[<span data-ttu-id="2437d-237">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="2437d-237">IdentityAPI</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)|

#### <a name="parameters"></a><span data-ttu-id="2437d-238">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2437d-238">Parameters</span></span>

<span data-ttu-id="2437d-239">`options`: Facultatif.</span><span class="sxs-lookup"><span data-stu-id="2437d-239">`options` - Optional.</span></span> <span data-ttu-id="2437d-240">Accepte un objet [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) (voir ci-dessous) pour définir les comportements d’authentification.</span><span class="sxs-lookup"><span data-stu-id="2437d-240">Accepts an [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="2437d-241">`callback`: Facultatif.</span><span class="sxs-lookup"><span data-stu-id="2437d-241">`callback` - Optional.</span></span> <span data-ttu-id="2437d-242">Accepte une méthode de rappel qui peut analyser le jeton pour l’ID de l’utilisateur ou utilisez le jeton dans le flux de « de la part de » pour accéder à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2437d-242">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="2437d-243">Si[AsyncResult](/javascript/api/office/office.asyncresult) `.status` est « Réussi », puis`AsyncResult.value` est le AAD v brut.</span><span class="sxs-lookup"><span data-stu-id="2437d-243">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="2437d-244">2.0: le jeton d’accès mis en forme.</span><span class="sxs-lookup"><span data-stu-id="2437d-244">2.0-formatted access token.</span></span>

<span data-ttu-id="2437d-245">L’interface [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) fournit des options pour l’expérience utilisateur quand Office reçoit un jeton d’accès pour le complément à partir d’AAD v.</span><span class="sxs-lookup"><span data-stu-id="2437d-245">The [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="2437d-246">2.0 avec la méthode`getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="2437d-246">2.0 with the `getAccessToken` method.</span></span>
