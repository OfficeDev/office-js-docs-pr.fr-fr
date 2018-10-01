---
title: Activer l’authentification unique pour des compléments Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 05b5088a61df3f77a09b60dbdc3129074d5f8530
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348169"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="ce412-102">Activer l’authentification unique pour des compléments Office (préversion)</span><span class="sxs-lookup"><span data-stu-id="ce412-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="ce412-103">Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365).</span><span class="sxs-lookup"><span data-stu-id="ce412-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="ce412-104">Vous pouvez profiter de cette fonctionnalité et utiliser l’authentification unique (SSO) pour autoriser l’utilisateur à accéder à votre complément sans qu’il ne soit obligé de se connecter une seconde fois.</span><span class="sxs-lookup"><span data-stu-id="ce412-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Image illustrant le processus de connexion pour un complément](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="ce412-106">État de la préversion</span><span class="sxs-lookup"><span data-stu-id="ce412-106">Preview Status</span></span>

<span data-ttu-id="ce412-107">L’API d'authentification unique est actuellement prise en charge uniquement en préversion.</span><span class="sxs-lookup"><span data-stu-id="ce412-107">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="ce412-108">Elle est disponible aux fins d'expérimentation pour les développeurs ; mais elle ne doit pas être utilisée dans un complément de production.</span><span class="sxs-lookup"><span data-stu-id="ce412-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="ce412-109">En outre, les compléments qui utilisent l’authentification unique ne sont pas acceptés dans [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="ce412-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="ce412-110">Certaines applications Office ne prennent pas en charge la préversion de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="ce412-110">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="ce412-111">Elle est disponible dans Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="ce412-111">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="ce412-112">Pour plus d’informations sur les applications qui prennent en charge actuellement l’API d’authentification unique, consultez la rubrique [Ensembles de conditions requises pour l’API d’identification](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="ce412-112">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="ce412-113">Configuration requise et meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="ce412-113">Requirements and Best Practices</span></span>

<span data-ttu-id="ce412-114">Pour utiliser l’authentification unique, vous devez charger la version bêta de la bibliothèque JavaScript de Office à partir de `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` au niveau de la page de démarrage HTML du complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="ce412-115">Si vous utilisez un complément **Outlook**, veillez à activer l’authentification moderne pour la location d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="ce412-115">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="ce412-116">Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="ce412-116">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="ce412-117">Vous ne devez *pas* utiliser l’authentification unique comme seule méthode d’authentification pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-117">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="ce412-118">Vous devez implémenter un système d’authentification alternatif que votre complément peut utiliser dans certaines situations d’erreur.</span><span class="sxs-lookup"><span data-stu-id="ce412-118">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="ce412-119">Vous pouvez utiliser un système de tables de l’utilisateur et d’authentification, ou vous pouvez tirer parti d’un des fournisseurs de connexion de mise en réseau.</span><span class="sxs-lookup"><span data-stu-id="ce412-119">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="ce412-120">Pour plus d’informations sur comment procéder avec un complément Office, voir [Autoriser les services externes dans votre complément Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ce412-120">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="ce412-121">Pour *Outlook*, il existe un système alternatif recommandé.</span><span class="sxs-lookup"><span data-stu-id="ce412-121">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="ce412-122">Pour plus d’informations, reportez-vous à [Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="ce412-122">For more details, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="ce412-123">Fonctionnement de l’authentification unique au moment de l’exécution</span><span class="sxs-lookup"><span data-stu-id="ce412-123">How it works at runtime</span></span>

<span data-ttu-id="ce412-124">Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="ce412-124">The following diagram shows how the SSO process works.</span></span>

![Diagramme illustrant le processus d’authentification unique](../images/sso-overview-diagram.png)

1. <span data-ttu-id="ce412-126">Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessTokenAsync](#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="ce412-126">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference).</span></span> <span data-ttu-id="ce412-127">Cela indique à l’application hôte Office qu’elle doit obtenir un jeton d’accès au complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-127">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="ce412-128">Voir [Exemple de token](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="ce412-128">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="ce412-129">Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.</span><span class="sxs-lookup"><span data-stu-id="ce412-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="ce412-130">Si c’est la première fois que l’utilisateur utilise votre complément, il est invité à donner son consentement.</span><span class="sxs-lookup"><span data-stu-id="ce412-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="ce412-131">L’application hôte Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="ce412-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="ce412-132">Azure AD envoie le jeton de complément à l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="ce412-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="ce412-133">L’application hôte Office envoie le **jeton de complément** au complément dans le cadre de l’objet de résultat renvoyé par l’appel `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="ce412-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="ce412-134">Dans le complément, JavaScript peut analyser le jeton et extraire les informations dont il a besoin, telles que l'adresse e-mail de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce412-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="ce412-135">Optionnellement, le complément peut envoyer une requête HTTP à son serveur pour obtenir plus de données sur l'utilisateur, notamment les préférences de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce412-135">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="ce412-136">Alternativement, le jeton d'accès lui-même pourrait être envoyé au serveur pour analyse et validation.</span><span class="sxs-lookup"><span data-stu-id="ce412-136">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="ce412-137">Développer un complément d’authentification unique</span><span class="sxs-lookup"><span data-stu-id="ce412-137">Develop an SSO add-in</span></span>

<span data-ttu-id="ce412-138">Cette section décrit les tâches que nécessite la création d’un complément Office qui utilise l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="ce412-138">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="ce412-139">Ces tâches sont décrites ici indépendamment du langage et de l’infrastructure.</span><span class="sxs-lookup"><span data-stu-id="ce412-139">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="ce412-140">Pour obtenir des exemples de procédures pas-à-pas détaillées, consultez les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="ce412-140">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="ce412-141">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="ce412-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="ce412-142">Créer un complément Office ASP.NET qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="ce412-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="ce412-143">Créer l’application de service</span><span class="sxs-lookup"><span data-stu-id="ce412-143">Create the service application</span></span>

<span data-ttu-id="ce412-144">Enregistrer le complément sur le portail d’inscription pour le point de terminaison Azure v2.0 : https://apps.dev.microsoft.com.</span><span class="sxs-lookup"><span data-stu-id="ce412-144">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="ce412-145">Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :</span><span class="sxs-lookup"><span data-stu-id="ce412-145">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="ce412-146">Obtenez un ID de client et un code secret pour le complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="ce412-147">Spécifiez les autorisations dont votre complément a besoin pour AAD v.</span><span class="sxs-lookup"><span data-stu-id="ce412-147">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="ce412-148">Endpoint 2.0 (et éventuellement Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="ce412-148">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="ce412-149">L'autorisation de "profil" est toujours nécessaire.</span><span class="sxs-lookup"><span data-stu-id="ce412-149">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="ce412-150">Accordez la confiance de l’application hôte Office au complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="ce412-151">Pré-autorisez l’application hôte Office pour le complément avec l’autorisation par défaut *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="ce412-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="ce412-152">Pour plus de détails sur ce processus, voir [Enregistrer un complément Office qui utilise l'authentification unique auprès du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="ce412-152">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="ce412-153">Configurez le complément</span><span class="sxs-lookup"><span data-stu-id="ce412-153">Configure the add-in</span></span>

<span data-ttu-id="ce412-154">Ajoutez un nouveau balisage au manifeste du complément :</span><span class="sxs-lookup"><span data-stu-id="ce412-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="ce412-155">**WebApplicationInfo** : parent des éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="ce412-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="ce412-156">**Id** - ID du client du complément : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-156">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="ce412-157">Voir [Enregistrer un complément Office qui utilise l'authentification unique (SSO) avec le point de terminaison AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="ce412-157">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="ce412-158">**Resource** : URL du complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="ce412-159">**Scopes** : parent d’un ou plusieurs éléments **Scope**.</span><span class="sxs-lookup"><span data-stu-id="ce412-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="ce412-160">**Scope** - Spécifie une autorisation dont le complément a besoin pour AAD.</span><span class="sxs-lookup"><span data-stu-id="ce412-160">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="ce412-161">L' `profile` autorisation est toujours nécessaire et il peut s'agir de la seule autorisation nécessaire si votre complément n'accède pas à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ce412-161">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="ce412-162">Si c'est le cas, vous avez également besoin des éléments d'une **étendue**pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="ce412-162">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="ce412-163">Les bibliothèques que vous utilisez dans votre code pour accéder à Microsoft Graph peuvent avoir besoin d'autorisations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="ce412-163">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="ce412-164">Par exemple, Microsoft Authentication Library (MSAL) pour .NET nécessite `offline_access` une autorisation.</span><span class="sxs-lookup"><span data-stu-id="ce412-164">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="ce412-165">Pour plus d'informations, voir [Autoriser Microsoft Graph à partir d'un complément Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="ce412-165">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="ce412-p113">Pour les hôtes Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="ce412-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="ce412-168">Voici un exemple de balise :</span><span class="sxs-lookup"><span data-stu-id="ce412-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="ce412-169">Ajouter du code côté client</span><span class="sxs-lookup"><span data-stu-id="ce412-169">Add client-side code</span></span>

<span data-ttu-id="ce412-170">Ajoutez un code JavaScript pour le complément à :</span><span class="sxs-lookup"><span data-stu-id="ce412-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="ce412-171">Appelez [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="ce412-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="ce412-172">Analyser le jeton ou le transmettre au code côté serveur du complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="ce412-173">Voici un exemple simple d'un appel à `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="ce412-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="ce412-174">Cet exemple ne présente explicitement qu'un seul type d'erreur.</span><span class="sxs-lookup"><span data-stu-id="ce412-174">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="ce412-175">Pour avoir des exemples de traitement des erreurs plus élaborés, voir [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) et [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="ce412-175">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="ce412-176">Voir également [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="ce412-176">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="ce412-177">Voici un exemple simple d’un passage de jeton du complément vers le serveur.</span><span class="sxs-lookup"><span data-stu-id="ce412-177">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="ce412-178">Le token est inclus en tant qu' `Authorization` en-tête lors de l'envoi d'une demande au serveur.</span><span class="sxs-lookup"><span data-stu-id="ce412-178">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="ce412-179">Dans cet exemple, l'envoi de données JSON se fait en utilisant la méthode `POST`, mais `GET` est suffisant pour envoyer le jeton d'accès lorsque vous n'écrivez pas sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="ce412-179">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="ce412-180">Quand appeler la méthode</span><span class="sxs-lookup"><span data-stu-id="ce412-180">When to call the method</span></span>

<span data-ttu-id="ce412-181">Si votre complément ne peut pas être utilisé lorsqu'aucun utilisateur n’est connecté à Office, vous devez appeler `getAccessTokenAsync` *au lancement du complément*.</span><span class="sxs-lookup"><span data-stu-id="ce412-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="ce412-182">Si le complément possède certaines fonctionnalités qui ne nécessitent pas d’utilisateur connecté, appelez `getAccessTokenAsync` *lorsque l’utilisateur effectue une action qui requiert un utilisateur connecté*.</span><span class="sxs-lookup"><span data-stu-id="ce412-182">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="ce412-183">Les appels répétés à `getAccessTokenAsync` ne causent aucune détérioration importante des performances, car Office met en cache le jeton d’accès et le réutilise jusqu'à ce qu’il arrive à expiration, sans effectuer un autre appel à l’AAD v.</span><span class="sxs-lookup"><span data-stu-id="ce412-183">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="ce412-184">2.0 dès que `getAccessTokenAsync` est appelé.</span><span class="sxs-lookup"><span data-stu-id="ce412-184">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="ce412-185">Ainsi, vous pouvez ajouter des appels de `getAccessTokenAsync` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le jeton est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="ce412-185">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="ce412-186">Ajouter du code côté serveur</span><span class="sxs-lookup"><span data-stu-id="ce412-186">Add server-side code</span></span>

<span data-ttu-id="ce412-187">Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement.</span><span class="sxs-lookup"><span data-stu-id="ce412-187">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="ce412-188">Quelques tâches côté serveur que votre complément pourrait faire :</span><span class="sxs-lookup"><span data-stu-id="ce412-188">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="ce412-189">Créer une ou plusieurs méthodes d'API Web qui utilisent des informations sur l'utilisateur qui sont extraites du jeton ; par exemple, une méthode qui recherche les préférences de l'utilisateur dans votre base de données hébergée.</span><span class="sxs-lookup"><span data-stu-id="ce412-189">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="ce412-190">(Voir **Utilisation du jeton SSO en tant qu'identité** ci-dessous). En fonction de votre langue et de votre structure, des bibliothèques peuvent être disponibles pour simplifier le code que vous devez écrire.</span><span class="sxs-lookup"><span data-stu-id="ce412-190">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="ce412-191">Obtenir des données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ce412-191">Get Microsoft Graph data.</span></span> <span data-ttu-id="ce412-192">Votre code côté serveur doit effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ce412-192">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="ce412-193">Valider le token (voir **Valider le token** ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="ce412-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="ce412-194">Démarrer le flux « de la part de » avec un appel du point de terminaison Azure AD v2.0 qui inclut le jeton d’accès du complément, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et code secret).</span><span class="sxs-lookup"><span data-stu-id="ce412-194">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="ce412-195">Dans ce contexte, le token est appelé token de démarrage.</span><span class="sxs-lookup"><span data-stu-id="ce412-195">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="ce412-196">Mettre en cache le nouveau jeton renvoyé par le flux intermédiaire.</span><span class="sxs-lookup"><span data-stu-id="ce412-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="ce412-197">Obtenir des données à partir de Microsoft Graph en utilisant le nouveau jeton.</span><span class="sxs-lookup"><span data-stu-id="ce412-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="ce412-198">Pour plus de détails sur l'obtention d'un accès autorisé aux données Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre complément Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="ce412-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="ce412-199">Valider le jeton d’accès</span><span class="sxs-lookup"><span data-stu-id="ce412-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="ce412-200">Quand l’API web reçoit le jeton d’accès, elle doit valider son fonctionnement avant de l’utiliser.</span><span class="sxs-lookup"><span data-stu-id="ce412-200">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="ce412-201">Le jeton est un jeton JWT (JSON Web Tokan). En d’autres termes, la validation se déroule comme dans la plupart des flux OAuth standard.</span><span class="sxs-lookup"><span data-stu-id="ce412-201">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="ce412-202">Il existe un certain nombre de bibliothèques pouvant gérer la validation JWT qui sont toutes, au minimum, chargées de :</span><span class="sxs-lookup"><span data-stu-id="ce412-202">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="ce412-203">vérifier que le jeton est bien formé ;</span><span class="sxs-lookup"><span data-stu-id="ce412-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="ce412-204">vérifier que le jeton a été émis par l’autorité souhaitée ;</span><span class="sxs-lookup"><span data-stu-id="ce412-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="ce412-205">vérifier que le jeton est destiné à l’API web.</span><span class="sxs-lookup"><span data-stu-id="ce412-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="ce412-206">Suivez les recommandations suivantes quand vous validez le jeton :</span><span class="sxs-lookup"><span data-stu-id="ce412-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="ce412-207">Les jetons SSO valides doivent être émis par l’autorité Azure `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="ce412-207">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="ce412-208">La revendication `iss` dans le jeton doit commencer par cette valeur.</span><span class="sxs-lookup"><span data-stu-id="ce412-208">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="ce412-209">Le paramètre `aud` du jeton devra correspondre à l’ID d’application de l’enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="ce412-210">Le paramètre `scp` du jeton devra correspondre à `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="ce412-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="ce412-211">Utilisation du jeton SSO comme identité</span><span class="sxs-lookup"><span data-stu-id="ce412-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="ce412-212">Si votre complément doit vérifier l’identité de l’utilisateur, le jeton SSO contient des informations utiles pour établir son identité.</span><span class="sxs-lookup"><span data-stu-id="ce412-212">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="ce412-213">Les lignes suivantes présentes dans le jeton concernent l’identité de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce412-213">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="ce412-214">`name` - Le nom de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce412-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="ce412-215">`preferred_username` -  L'adresse e-mail de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ce412-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="ce412-216">`oid` - Un GUID représentant l’ID de l’utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="ce412-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="ce412-217">`tid` - Un GUID représentant l'ID de l'organisation de l'utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="ce412-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="ce412-218">Étant donné que les valeurs `name` et `preferred_username` peuvent être amenées à changer, nous vous recommandons d’utiliser les valeurs `oid` et `tid` pour corréler l’identité de l’utilisateur avec le service d’autorisation de votre API principale.</span><span class="sxs-lookup"><span data-stu-id="ce412-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="ce412-219">Par exemple, votre service peut mettre en forme ces valeurs de la façon suivante `{oid-value}@{tid-value}`, puis stocker cette mise en forme sous forme de valeur dans l’enregistrement de l’utilisateur dans votre base de données utilisateur interne.</span><span class="sxs-lookup"><span data-stu-id="ce412-219">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="ce412-220">Puis à l'aide de requêtes ultérieures, l’utilisateur pourra être récupéré grâce à cette valeur et l’accès à certaines ressources pourra être déterminé selon les mécanismes de contrôle d’accès existants.</span><span class="sxs-lookup"><span data-stu-id="ce412-220">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="ce412-221">Exemple de jeton</span><span class="sxs-lookup"><span data-stu-id="ce412-221">Example access token</span></span>

<span data-ttu-id="ce412-222">Voici une charge utile décodée typique de jeton.</span><span class="sxs-lookup"><span data-stu-id="ce412-222">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="ce412-223">Pour plus d'informations sur les propriétés, voir [Référence des jetons Azure Active Directory v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="ce412-223">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="ce412-224">Utilisation de l'authentification unique (SSO) avec un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="ce412-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="ce412-225">Il existe quelques différences mineures, mais importantes, en ce qui concerne l'utilisation de la connexion unique SSO dans et comme complément Outlook comparé à son utilisation comme complément Excel, PowerPoint ou Word.</span><span class="sxs-lookup"><span data-stu-id="ce412-225">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="ce412-226">Assurez-vous de lire [Authentifier un utilisateur avec un jeton unique logé dans le complément Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) et l' [étude de cas : Implémenter la connexion unique à votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="ce412-226">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="ce412-227">Référence de l’API de l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="ce412-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="ce412-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ce412-228">getAccessTokenAsync</span></span>

<span data-ttu-id="ce412-229">L’espace de noms Office Auth, `Office.context.auth`, fournit une méthode, `getAccessTokenAsync` qui permet à l’hôte Office d'obtenir un jeton d’accès à l’application web du module complémentaire.</span><span class="sxs-lookup"><span data-stu-id="ce412-229">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="ce412-230">Indirectement, cela permet également au module complémentaire d'accéder aux données Microsoft Graph de l'utilisateur connecté sans qu'il ait à se connecter une seconde fois.</span><span class="sxs-lookup"><span data-stu-id="ce412-230">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="ce412-231">La méthode appelle le point de terminaison Azure Active Directory V 2.0 pour obtenir un jeton d'accès à l'application Web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="ce412-231">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="ce412-232">Cela permet aux compléments d’identifier les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="ce412-232">This enables add-ins to identify users.</span></span> <span data-ttu-id="ce412-233">Le Code côté serveur peut utiliser ce jeton pour accéder à Microsoft Graph pour l’application web du complément à l’aide du [flux OAuth intermédiaire](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="ce412-233">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="ce412-234">Dans Outlook, cette API n'est pas prise en charge si le complément est chargé dans une boîte aux lettres Outlook.com ou Gmail.</span><span class="sxs-lookup"><span data-stu-id="ce412-234">[!Note In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.]</span></span>

<table><tr><td><span data-ttu-id="ce412-235">Hôtes</span><span class="sxs-lookup"><span data-stu-id="ce412-235">Hosts</span></span></td><td><span data-ttu-id="ce412-236">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="ce412-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="ce412-237">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ce412-237">Requirement sets</span></span></td><td>[<span data-ttu-id="ce412-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="ce412-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="ce412-239">Paramètres</span><span class="sxs-lookup"><span data-stu-id="ce412-239">Parameters</span></span>

<span data-ttu-id="ce412-240">`options` - Facultatif.</span><span class="sxs-lookup"><span data-stu-id="ce412-240">`options` - Optional.</span></span> <span data-ttu-id="ce412-241">Accepte un objet `AuthOptions` (voir ci-dessous) pour définir les comportements de connexion.</span><span class="sxs-lookup"><span data-stu-id="ce412-241">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="ce412-242">`callback` - Facultatif.</span><span class="sxs-lookup"><span data-stu-id="ce412-242">`callback` - Optional.</span></span> <span data-ttu-id="ce412-243">Accepte une méthode de rappel qui peut analyser le jeton pour l’ID d’utilisateur ou utiliser le jeton dans le flux « de la part de » pour accéder à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ce412-243">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="ce412-244">Si [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status`  est « réussi », alors `AsyncResult.value` est l'AAD v brut.</span><span class="sxs-lookup"><span data-stu-id="ce412-244">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="ce412-245">Jeton d'accès au format 2.0.</span><span class="sxs-lookup"><span data-stu-id="ce412-245">2.0-formatted access token.</span></span>

<span data-ttu-id="ce412-246">L' `AuthOptions` interface fournit des options pour l'expérience utilisateur, lorsque Office obtient un jeton d'accès au complément d'AAD v.</span><span class="sxs-lookup"><span data-stu-id="ce412-246">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="ce412-247">2.0 avec la méthode `getAccessTokenAsync` .</span><span class="sxs-lookup"><span data-stu-id="ce412-247">2.0 with the `getAccessTokenAsync` method.</span></span>

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



