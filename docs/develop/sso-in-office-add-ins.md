---
title: Activer l’authentification unique pour des compléments Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: ce57c5d70e2c48a89b2fd84c30ac7b8580650896
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876605"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="0a326-102">Activer l’authentification unique pour des compléments Office (aperçu)</span><span class="sxs-lookup"><span data-stu-id="0a326-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="0a326-103">Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365).</span><span class="sxs-lookup"><span data-stu-id="0a326-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="0a326-104">Vous pouvez profiter de cette fonctionnalité et utiliser l’authentification unique (SSO) pour autoriser l’utilisateur à accéder à votre complément sans qu’il ne soit obligé de se connecter une seconde fois.</span><span class="sxs-lookup"><span data-stu-id="0a326-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>


![Image illustrant le processus de connexion pour un complément](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> <span data-ttu-id="0a326-106">L’API de l’authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="0a326-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="0a326-107">Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles d'exigences de l’API d’identité](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="0a326-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span></span>
> <span data-ttu-id="0a326-108">Pour utiliser l’authentification unique, vous devez charger la version bêta de la bibliothèque JavaScript d'Office à partir de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js dans la page de démarrage HTML du complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-108">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>
> <span data-ttu-id="0a326-109">Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour le client Office 365.</span><span class="sxs-lookup"><span data-stu-id="0a326-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="0a326-110">Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="0a326-110">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="0a326-111">Pour les utilisateurs, cela permet une exécution aisée de votre complément qui ne requiert qu’une seule connexion.</span><span class="sxs-lookup"><span data-stu-id="0a326-111">For users, this makes running your add-in a smooth experience that involves at signing in only once.</span></span> <span data-ttu-id="0a326-112">Pour les développeurs, cela signifie que votre complément n'a pas besoin de gérer ses propres tables utilisateur avec des mots de passe cryptés.</span><span class="sxs-lookup"><span data-stu-id="0a326-112">For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.</span></span>

### <a name="how-it-works-at-runtime"></a><span data-ttu-id="0a326-113">Mode de fonctionnement en cours d’exécution</span><span class="sxs-lookup"><span data-stu-id="0a326-113">How it works at runtime</span></span>

<span data-ttu-id="0a326-114">Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="0a326-114">The following diagram shows how the SSO process works.</span></span>

![Diagramme illustrant le processus d’authentification unique](../images/sso-overview-diagram.png)

1. <span data-ttu-id="0a326-p104">Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessTokenAsync](#sso-api-reference). Cela indique à l'application hôte Office d'obtenir un jeton d'accès au complément. Voir [Exemple de jeton d'accès](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="0a326-p104">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. (Hereafter, this is called the [add-in token](#example-access-token).)</span></span>
2. <span data-ttu-id="0a326-119">Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.</span><span class="sxs-lookup"><span data-stu-id="0a326-119">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="0a326-120">Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.</span><span class="sxs-lookup"><span data-stu-id="0a326-120">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="0a326-121">L’application hôte Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="0a326-121">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="0a326-122">Azure AD envoie le jeton de complément à l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="0a326-122">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="0a326-123">L’application hôte Office envoie le **jeton de complément** au complémentent en tant que partie de l’objet de résultat renvoyé par `getAccessTokenAsync` l’appel.</span><span class="sxs-lookup"><span data-stu-id="0a326-123">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="0a326-124">Dans le complément, JavaScript peut analyser le token et extraire les informations dont il a besoin, telles que l'adresse e-mail de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0a326-124">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="0a326-125">Optionnellement, le complément peut envoyer une requête HTTP à son serveur pour obtenir plus de données sur l'utilisateur, notamment les préférences de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0a326-125">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="0a326-126">Alternativement, le token lui-même pourrait être envoyé au serveur pour analyse et validation.</span><span class="sxs-lookup"><span data-stu-id="0a326-126">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="0a326-127">Développer un complément SSO</span><span class="sxs-lookup"><span data-stu-id="0a326-127">Develop an SSO add-in</span></span>

<span data-ttu-id="0a326-128">Cette section décrit les tâches impliquées dans la création d’un complément Office qui utilise l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="0a326-128">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="0a326-129">Ces tâches sont décrites ici indépendamment du langage et de l’infrastructure.</span><span class="sxs-lookup"><span data-stu-id="0a326-129">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="0a326-130">Pour obtenir des exemples de procédures pas à pas détaillées, consultez les rubriques suivantes :</span><span class="sxs-lookup"><span data-stu-id="0a326-130">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="0a326-131">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="0a326-131">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="0a326-132">Créer un complément Office ASP.NET qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="0a326-132">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="0a326-133">Créer l’application de service</span><span class="sxs-lookup"><span data-stu-id="0a326-133">Create the service application</span></span>

<span data-ttu-id="0a326-134">Enregistrer le complément sur le portail d’inscription pour le point de terminaison Azure v2.0 : https://apps.dev.microsoft.com.</span><span class="sxs-lookup"><span data-stu-id="0a326-134">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="0a326-135">Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :</span><span class="sxs-lookup"><span data-stu-id="0a326-135">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="0a326-136">Obtenez un ID de client et un code secret pour le complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-136">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="0a326-137">Spécifiez les autorisations dont votre complément a besoin pour AAD v.</span><span class="sxs-lookup"><span data-stu-id="0a326-137">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="0a326-138">Point de terminaison 2.0 (et éventuellement Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="0a326-138">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="0a326-139">L'autorisation "profil" est toujours nécessaire.</span><span class="sxs-lookup"><span data-stu-id="0a326-139">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="0a326-140">Accordez la confiance de l’application hôte Office au complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-140">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="0a326-141">Pré-autorisez l’application hôte Office pour le complément avec l’autorisation par défaut *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="0a326-141">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="0a326-142">Pour plus de détails sur ce processus, voir [Enregistrer un complément Office qui utilise l'authentification unique auprès du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="0a326-142">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="0a326-143">Configurer le complément</span><span class="sxs-lookup"><span data-stu-id="0a326-143">Configure the add-in</span></span>

<span data-ttu-id="0a326-144">Ajoutez un nouveau balisage au manifeste du complément :</span><span class="sxs-lookup"><span data-stu-id="0a326-144">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="0a326-145">**WebApplicationInfo** : parent des éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="0a326-145">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="0a326-146">**Id** - ID du client du complément : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-146">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="0a326-147">Voir [Enregistrer un complément Office qui utilise l'authentification unique (SSO) avec le point de terminaison AD v2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="0a326-147">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="0a326-148">**Resource** -  L'URL du complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-148">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="0a326-149">**Étendues** : Le parent d’une ou plusieurs éléments **Étendue**.</span><span class="sxs-lookup"><span data-stu-id="0a326-149">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="0a326-150">**Étendue** - Spécifie une autorisation dont le complément a besoin pour AAD.</span><span class="sxs-lookup"><span data-stu-id="0a326-150">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="0a326-151">L' `profile` autorisation est toujours nécessaire et il peut s'agir de la seule autorisation nécessaire si votre complément n'accède pas à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="0a326-151">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="0a326-152">Si c'est le cas, vous avez également besoin des éléments d'une **étendue**pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="0a326-152">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="0a326-153">Les bibliothèques que vous utilisez dans votre code pour accéder à Microsoft Graph peuvent avoir des besoin d'autorisations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="0a326-153">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="0a326-154">Par exemple, Microsoft Authentication Library (MSAL) pour .NET nécessite `offline_access` une autorisation.</span><span class="sxs-lookup"><span data-stu-id="0a326-154">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="0a326-155">Pour plus d'informations, voir [Autoriser Microsoft Graph à partir d'un complément Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="0a326-155">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="0a326-p111">Pour les hôtes Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="0a326-p111">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="0a326-158">Voici un exemple de balise :</span><span class="sxs-lookup"><span data-stu-id="0a326-158">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="0a326-159">Ajouter du code côté client</span><span class="sxs-lookup"><span data-stu-id="0a326-159">Add client-side code</span></span>

<span data-ttu-id="0a326-160">Ajoutez un code JavaScript pour le complément à :</span><span class="sxs-lookup"><span data-stu-id="0a326-160">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="0a326-161">Appeler [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span><span class="sxs-lookup"><span data-stu-id="0a326-161">Call [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span></span>
* <span data-ttu-id="0a326-162">Analyser le token ou le transmettre au code côté serveur du complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-162">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="0a326-163">Voici un exemple simple d'un appel à `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="0a326-163">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!Note]
> <span data-ttu-id="0a326-164">Cet exemple ne présente explicitement qu'un seul type d'erreur.</span><span class="sxs-lookup"><span data-stu-id="0a326-164">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="0a326-165">Pour avoir des exemples de traitement des erreurs plus élaborés, voir [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) et [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="0a326-165">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="0a326-166">Voir également [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="0a326-166">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="0a326-167">Voici un exemple simple d’un passage de token du complément vers le serveur.</span><span class="sxs-lookup"><span data-stu-id="0a326-167">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="0a326-168">Le token est inclus en tant qu' `Authorization` en-tête lors de l'envoi d'une demande au serveur.</span><span class="sxs-lookup"><span data-stu-id="0a326-168">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="0a326-169">Dans cet exemple, l'envoi de données JSON se fait en utilisant la méthode `POST`, mais `GET` est suffisant pour envoyer le token d'accès lorsque vous n'écrivez pas sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="0a326-169">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="0a326-170">Quand appeler la méthode</span><span class="sxs-lookup"><span data-stu-id="0a326-170">When to call the method</span></span>

<span data-ttu-id="0a326-171">Si votre complément ne peut pas être utilisé lorsque aucun utilisateur n’est connecté à Office, vous devez appeler `getAccessTokenAsync` *au lancement du complément*.</span><span class="sxs-lookup"><span data-stu-id="0a326-171">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="0a326-172">Si le complément possède certaines fonctionnalités qui ne nécessitent pas d’utilisateur connecté, appelez `getAccessTokenAsync` *lorsque l’utilisateur effectue une action qui requiert un utilisateur connecté*.</span><span class="sxs-lookup"><span data-stu-id="0a326-172">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="0a326-173">Les appels répétés à `getAccessTokenAsync` ne causent aucune détérioration importante des performances, car Office met en cache le jeton d’accès et le réutilise jusqu'à ce qu’il arrive à expiration, sans effectuer un autre appel à l’AAD v.</span><span class="sxs-lookup"><span data-stu-id="0a326-173">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="0a326-174">2.0 dès que `getAccessTokenAsync` est appelé.</span><span class="sxs-lookup"><span data-stu-id="0a326-174">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="0a326-175">Ainsi, vous pouvez ajouter des appels de `getAccessTokenAsync` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le jeton est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="0a326-175">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="0a326-176">Ajouter du code côté serveur</span><span class="sxs-lookup"><span data-stu-id="0a326-176">Add server-side code</span></span>

<span data-ttu-id="0a326-177">Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement.</span><span class="sxs-lookup"><span data-stu-id="0a326-177">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="0a326-178">Quelques tâches côté serveur que votre complément pourrait faire :</span><span class="sxs-lookup"><span data-stu-id="0a326-178">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="0a326-179">Créer d'une ou plusieurs méthodes d'API Web qui utilisent des informations sur l'utilisateur qui sont extraitent du token ; par exemple, une méthode qui recherche les préférences de l'utilisateur dans votre base de données hébergée.</span><span class="sxs-lookup"><span data-stu-id="0a326-179">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="0a326-180">(Voir **Utilisation du token SSO en tant qu'identité** ci-dessous). En fonction de votre langue et de votre structure, des bibliothèques peuvent être disponibles pour simplifier le code que vous devez écrire.</span><span class="sxs-lookup"><span data-stu-id="0a326-180">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="0a326-181">Obtenir des données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="0a326-181">Get Microsoft Graph data.</span></span> <span data-ttu-id="0a326-182">Votre code côté serveur doit effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="0a326-182">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="0a326-183">Valider le token (voir **Valider le token** ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="0a326-183">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="0a326-184">Démarrer le flux « de la part de » avec un appel du point de terminaison Azure AD v2.0 qui inclut le jeton d’accès du complément, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et code secret).</span><span class="sxs-lookup"><span data-stu-id="0a326-184">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="0a326-185">Dans ce contexte, le token est appelé token de démarrage.</span><span class="sxs-lookup"><span data-stu-id="0a326-185">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="0a326-186">Mettre en cache le nouveau token renvoyé par le flux « de la part de ».</span><span class="sxs-lookup"><span data-stu-id="0a326-186">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="0a326-187">Obtenir des données à partir de Microsoft Graph en utilisant le nouveau jeton.</span><span class="sxs-lookup"><span data-stu-id="0a326-187">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="0a326-188">Pour plus de détails sur l'obtention d'un accès autorisé aux données Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre complément Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="0a326-188">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="0a326-189">Valider le jeton d’accès</span><span class="sxs-lookup"><span data-stu-id="0a326-189">For more information, see Validate the access token.</span></span>

<span data-ttu-id="0a326-190">Quand l’API web reçoit le jeton d’accès, elle doit le valider avant de l'utiliser.</span><span class="sxs-lookup"><span data-stu-id="0a326-190">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="0a326-191">Le jeton est un jeton JWT. En d’autres termes, la validation se déroule comme dans la plupart des flux OAuth standard.</span><span class="sxs-lookup"><span data-stu-id="0a326-191">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="0a326-192">Il existe un certain nombre de bibliothèques pouvant gérer la validation JWT qui sont toutes, au minimum, chargées de :</span><span class="sxs-lookup"><span data-stu-id="0a326-192">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="0a326-193">vérifier que le jeton est bien formé ;</span><span class="sxs-lookup"><span data-stu-id="0a326-193">Checking that the token is well-formed</span></span>
- <span data-ttu-id="0a326-194">vérifier que le jeton a été émis par l’autorité souhaitée ;</span><span class="sxs-lookup"><span data-stu-id="0a326-194">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="0a326-195">vérifier que le jeton est destiné à l’API web.</span><span class="sxs-lookup"><span data-stu-id="0a326-195">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="0a326-196">Suivez les recommandations suivantes quand vous validez le jeton :</span><span class="sxs-lookup"><span data-stu-id="0a326-196">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="0a326-197">Les jetons SSO valides doivent être émis par l’autorité Azure `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="0a326-197">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="0a326-198">La revendication `iss` dans le jeton doit commencer par cette valeur.</span><span class="sxs-lookup"><span data-stu-id="0a326-198">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="0a326-199">Le paramètre `aud` du jeton devra correspondre à l’ID d’application de l’enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-199">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="0a326-200">Le paramètre `scp` du jeton devra correspondre à `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="0a326-200">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="0a326-201">Utilisation du jeton SSO comme identité</span><span class="sxs-lookup"><span data-stu-id="0a326-201">Using the SSO token as an identity</span></span>

<span data-ttu-id="0a326-202">Si votre complément doit vérifier l’identité de l’utilisateur, le jeton SSO contient des informations utiles pour établir son identité.</span><span class="sxs-lookup"><span data-stu-id="0a326-202">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="0a326-203">Les revendications suivantes présentes dans le jeton concernent l’identité de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0a326-203">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="0a326-204">`name` - Le nom d'affichage de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0a326-204">`name` - The user's display name.</span></span>
- <span data-ttu-id="0a326-205">`preferred_username` -  L'adresse e-mail de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0a326-205">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="0a326-206">`oid` - Un GUID représentant l'ID de l'utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="0a326-206">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="0a326-207">`tid` - Un GUID représentant l'ID de l'organisation de l'utilisateur dans Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="0a326-207">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="0a326-208">Étant donné que les valeurs `name` et `preferred_username` peuvent être amenées à changer, nous vous recommandons d’utiliser les valeurs `oid` et `tid` pour corréler l’identité de l’utilisateur avec le service d’autorisation de votre API principale.</span><span class="sxs-lookup"><span data-stu-id="0a326-208">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="0a326-209">Par exemple, votre service peut mettre en forme ces valeurs de la façon suivante `{oid-value}@{tid-value}`, puis stocker cette mise en forme sous forme de valeur dans l’enregistrement de l’utilisateur dans votre base de données utilisateur interne.</span><span class="sxs-lookup"><span data-stu-id="0a326-209">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="0a326-210">Lors des demandes ultérieures, l’utilisateur pourra être récupéré grâce à cette valeur et l’accès à certaines ressources pourra être déterminé selon les mécanismes de contrôle d’accès existants.</span><span class="sxs-lookup"><span data-stu-id="0a326-210">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="0a326-211">Exemple de token</span><span class="sxs-lookup"><span data-stu-id="0a326-211">Example access token</span></span>

<span data-ttu-id="0a326-212">Voici une charge utile décodée typique de token.</span><span class="sxs-lookup"><span data-stu-id="0a326-212">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="0a326-213">Pour plus d'informations sur les propriétés, voir [Référence des jetons Azure Active Directory v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="0a326-213">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="0a326-214">Utilisation de l'authentification unique (SSO) avec un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="0a326-214">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="0a326-p124">Il y a quelques petites, mais importantes différences entre l'utilisation du SSO dans un complément Outlook et son utilisation dans un complément Excel, PowerPoint ou Word. Veillez lire [Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) et [Scénario : Implémenter la connexion unique à votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="0a326-p124">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="0a326-217">Référence de l’API de l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="0a326-217">SSO API reference</span></span>

<span data-ttu-id="0a326-218">L’espace de noms Office Auth, `Office.context.auth`, fournit une méthode, `getAccessTokenAsync` qui permet à l’hôte Office d'obtenir un jeton d’accès à l’application web du module complémentaire.</span><span class="sxs-lookup"><span data-stu-id="0a326-218">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="0a326-219">Indirectement, cela permet également au module complémentaire d'accéder aux données Microsoft Graph de l'utilisateur connecté sans qu'il ait à se connecter une seconde fois.</span><span class="sxs-lookup"><span data-stu-id="0a326-219">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="0a326-220">La méthode appelle le point de terminaison Azure Active Directory V 2.0 pour obtenir un jeton d'accès à l'application Web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="0a326-220">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="0a326-221">Cela permet aux compléments d’identifier les utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="0a326-221">This enables add-ins to identify users.</span></span> <span data-ttu-id="0a326-222">Le Code côté serveur peut utiliser ce jeton pour accéder à Microsoft Graph pour l’application web du complément à l’aide du [flux OAuth « Pour le compte de »](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="0a326-222">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [! Remarque : Dans Outlook, cette API n'est pas prise en charge si le complément est chargé dans une boîte aux lettres Outlook.com ou Gmail. ]

<table><tr><td><span data-ttu-id="0a326-224">Hôtes</span><span class="sxs-lookup"><span data-stu-id="0a326-224">Hosts</span></span></td><td><span data-ttu-id="0a326-225">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="0a326-225">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="0a326-226">Ensembles d'exigences</span><span class="sxs-lookup"><span data-stu-id="0a326-226">Requirement sets</span></span></td><td>[<span data-ttu-id="0a326-227">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="0a326-227">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

### <a name="parameters"></a><span data-ttu-id="0a326-228">Paramètres</span><span class="sxs-lookup"><span data-stu-id="0a326-228">Parameters</span></span>

<span data-ttu-id="0a326-229">`options` - Facultatif.</span><span class="sxs-lookup"><span data-stu-id="0a326-229">`options` - Optional.</span></span> <span data-ttu-id="0a326-230">Accepte un objet `AuthOptions` (voir ci-dessous) pour définir les comportements de connexion.</span><span class="sxs-lookup"><span data-stu-id="0a326-230">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="0a326-231">`callback` - Facultatif.</span><span class="sxs-lookup"><span data-stu-id="0a326-231">`callback` - Optional.</span></span> <span data-ttu-id="0a326-232">Accepte une méthode de rappel qui peut analyser le jeton pour l’ID d’utilisateur ou utiliser le jeton dans le flux « de la part de » pour accéder à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="0a326-232">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="0a326-233">Si [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status`  est « réussi », alors `AsyncResult.value` est l'AAD v brut.</span><span class="sxs-lookup"><span data-stu-id="0a326-233">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="0a326-234">Jeton d'accès au format 2.0.</span><span class="sxs-lookup"><span data-stu-id="0a326-234">2.0-formatted access token.</span></span>

<span data-ttu-id="0a326-235">L' `AuthOptions` interface fournit des options pour l'expérience utilisateur, lorsque Office obtient un jeton d'accès au complément d'AAD v.</span><span class="sxs-lookup"><span data-stu-id="0a326-235">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="0a326-236">2.0 avec la méthode `getAccessTokenAsync` .</span><span class="sxs-lookup"><span data-stu-id="0a326-236">2.0 with the `getAccessTokenAsync` method.</span></span>

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



