---
title: Création d’un complément Office Node.js qui utilise l’authentification unique
description: Apprenez à créer un complément basé sur Node.js utilisant l’authentification unique Office.
ms.date: 11/20/2019
localization_priority: Priority
ms.openlocfilehash: 362ca4a534800a683284b049e6e53776b1aa7f38
ms.sourcegitcommit: 013886c1b08ef2b378cf80bb88bc73ec56c3e869
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/22/2019
ms.locfileid: "39191738"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="b9bb1-103">Créer un complément Office Node.js qui utilise l’authentification unique (aperçu)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-103">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="b9bb1-p101">Les utilisateurs peuvent se connecter à Office et votre complément Web Office peut tirer parti de cette procédure de connexion pour autoriser les utilisateurs à accéder à votre complément et à Microsoft Graph sans obliger les utilisateurs à se connecter une deuxième fois. Pour obtenir une vue d’ensemble, consultez [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="b9bb1-106">Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré à Node.js et Express.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="b9bb1-107">Pour voir un article similaire sur un complément basé sur ASP.NET, reportez-vous à [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b9bb1-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="b9bb1-108">Prerequisites</span></span>

* <span data-ttu-id="b9bb1-109">[Nœud et npm](https://nodejs.org/), version 10.15.0 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-109">[Node and npm](https://nodejs.org/), version 10.15.0 or later.</span></span>

* <span data-ttu-id="b9bb1-110">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-110">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="b9bb1-111">TypeScript version 3.6.2 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-111">TypeScript, version 3.6.2 or later</span></span>

* <span data-ttu-id="b9bb1-112">Compte Office 365 (version abonnement d’Office) que vous pouvez obtenir en rejoignant le [programme pour les développeurs Office 365](https://aka.ms/devprogramsignup) et qui inclut un abonnement gratuit de 1 an à Office 365.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-112">An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.</span></span> <span data-ttu-id="b9bb1-113">Nous vous recommandons d’utiliser la version mensuelle la plus récente et la build du canal Office Insider, mais vous devez être un participant au programme Office Insider pour l’obtenir.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-113">You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version.</span></span> <span data-ttu-id="b9bb1-114">Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-114">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="b9bb1-115">Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-115">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

* <span data-ttu-id="b9bb1-116">Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-116">A source code editor.</span></span> <span data-ttu-id="b9bb1-117">Nous vous recommandons Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-117">We recommend Visual Studio Code.</span></span>

* <span data-ttu-id="b9bb1-118">Au moins des fichiers et classeurs sont stockés sur OneDrive Entreprise dans votre abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-118">At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.</span></span>

* <span data-ttu-id="b9bb1-119">Un locataire Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-119">A Microsoft Azure Tenant.</span></span> <span data-ttu-id="b9bb1-120">Ce complément requiert Azure Active Directory (AD).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-120">This add-in requires Azure Active Directiory (AD).</span></span> <span data-ttu-id="b9bb1-121">Azure AD fournit des services d’identité que les applications utilisent à des fins d’authentification et d’autorisation.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-121">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="b9bb1-122">Un abonnement d’évaluation peut être obtenu sur le site de [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-122">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="b9bb1-123">Configurer le projet de démarrage</span><span class="sxs-lookup"><span data-stu-id="b9bb1-123">Set up the starter project</span></span>

1. <span data-ttu-id="b9bb1-124">Clonez ou téléchargez le référentiel sur [Complément Office NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-124">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="b9bb1-125">Il existe trois versions de l’échantillon :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-125">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="b9bb1-p105">Le dossier **Before** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-p105">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="b9bb1-129">La version **Complète** de l’échantillon s’apparente au complément obtenu si vous aviez terminé les procédures de cet article, sauf que le projet final comporte des commentaires de code qui seraient redondants avec le texte de cet article.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-129">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="b9bb1-130">Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et Code côté serveur.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-130">To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="b9bb1-131">La version **SSOAutoSetup** est un exemple complet qui permet d’automatiser la plupart des étapes d’inscription du complément avec Azure AD et sa configuration.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-131">The **SSOAutoSetup** version is a completed sample that automates most of the steps to register the add-in with Azure AD and configure it.</span></span> <span data-ttu-id="b9bb1-132">Utilisez cette version si vous voulez rapidement afficher un complément opérationnel avec SSO.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-132">Use this version if you want to see a working add-in with SSO quickly.</span></span> <span data-ttu-id="b9bb1-133">Suivez simplement les étapes décrites dans le fichier Lisez-moi du dossier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-133">Just follow the steps in the Readme of the folder.</span></span> <span data-ttu-id="b9bb1-134">Nous vous recommandons, à un certain point, de suivre les étapes d’inscription et de configuration manuelles décrites dans cet article pour mieux comprendre la relation entre Azure AD et un complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-134">We recommend that at some point you go through the manual registration and setup steps in this article to better understand the relationship between Azure AD and an add-in.</span></span> 


1. <span data-ttu-id="b9bb1-135">Ouvrez une invite de commandes dans le dossier **auparavant**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-135">Open a command prompt in the **Before** folder.</span></span>

1. <span data-ttu-id="b9bb1-136">Saisissez `npm install`dans la console pour installer toutes les dépendances détaillées dans le fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-136">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="b9bb1-137">Exécutez la commande `npm run install-dev-certs`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-137">Run the command  `npm run install-dev-certs`.</span></span> <span data-ttu-id="b9bb1-138">Sélectionnez **Oui** lorsque vous êtes invité à installer le certificat.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-138">Select **Yes** to the prompt to disable the designer.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="b9bb1-139">Enregistrez le complément avec le point de terminaison Azure AD v2.0</span><span class="sxs-lookup"><span data-stu-id="b9bb1-139">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="b9bb1-140">Accédez à la page [portail Azure : enregistrement des applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour enregistrer votre application.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-140">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="b9bb1-141">Connectez-vous à votre client Office 365 en utilisant les informations d’identification d’***administrateur***.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-141">Sign in with the ***admin*** credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="b9bb1-142">Par exemple, MonNom@contoso.onmicrosoft.com.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-142">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="b9bb1-143">Sélectionnez **Nouvelle inscription**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-143">Select **New registration**.</span></span> <span data-ttu-id="b9bb1-144">Sur la page **Inscrire une application**, définissez les valeurs comme suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-144">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="b9bb1-145">Définissez le **Nom** sur `Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-145">Set **Name** to `Office-Add-in-NodeJS-SSO`.</span></span>
    * <span data-ttu-id="b9bb1-146">Définissez les **Types de comptes pris en charge** à **Comptes dans un annuaire organisationnel et les comptes personnels Microsoft (par ex. Skype, Xbox et Outlook.com)**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-146">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span>
    * <span data-ttu-id="b9bb1-147">Configurez **URI de redirection** vers` https://localhost:44355/dialog.html`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-147">Set **Redirect URI** to` https://localhost:44355/dialog.html`.</span></span>
    * <span data-ttu-id="b9bb1-148">Choisissez **Inscrire**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-148">Choose **Register**.</span></span>

1. <span data-ttu-id="b9bb1-149">Sur la page **Office-Add-in-NodeJS-SSO**, copiez et enregistrez les valeurs pour l’**ID de l’application (client)** et l’**ID de répertoire (client)**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-149">On the **$ADD-IN-NAME$** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="b9bb1-150">Vous utiliserez les deux plus tard.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-150">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b9bb1-151">Cet ID a la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) demandent un accès autorisé à l’application.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-151">This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="b9bb1-152">Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-152">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="b9bb1-153">Sous **Gérer**, sélectionnez **Authentification**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-153">Select **Authentication** under **Manage**.</span></span> <span data-ttu-id="b9bb1-154">Dans la section **Implict Grant**, activez les cases à cocher pour **Jeton d’accès** et **Jeton d’ID**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-154">In the **Implict grant** section, enable the checkboxes for both **Access token** and **ID token**.</span></span> <span data-ttu-id="b9bb1-155">L’exemple dispose d’un système d’autorisation de secours qui est appelé lorsque l’authentification unique n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-155">The sample has a fallback authorization system that is invoked when SSO is not available.</span></span> <span data-ttu-id="b9bb1-156">Le système utilise le Flux implicite.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-156">This system uses the Implicit Flow.</span></span>

1. <span data-ttu-id="b9bb1-157">Sélectionnez **Enregistrer** en haut du formulaire.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-157">Select **Save** at the top of the form.</span></span>

1. <span data-ttu-id="b9bb1-158">Sélectionnez **Certificats et secrets** sous **Gérer**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-158">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="b9bb1-159">Sélectionnez le bouton **Nouveau secret client**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-159">Select the **New client secret** button.</span></span> <span data-ttu-id="b9bb1-160">Entrer une valeur pour **Description** puis sélectionnez une option appropriée pour **Expire le** puis **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-160">Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="b9bb1-161">*Copier la valeur secrète client immédiatement et enregistrez-la avec l’ID d’application* avant de continuer car vous en aurez besoin dans une procédure plus loin.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-161">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="b9bb1-162">Sélectionnez **Exposer une API** sous **Gérer**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-162">Select **Expose an API** under **Manage**.</span></span> <span data-ttu-id="b9bb1-163">Sélectionnez le lien **Définir** pour générer l’URI de l’ID d’application sous la forme « api://$App ID GUID$ », où $App ID GUID$ est l’**ID de l’application (client)**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-163">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="b9bb1-164">Insérez `localhost:44355/` (remarquez la barre oblique « / » ajoutée à la fin) entre les doubles barres obliques et le GUID.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-164">Insert the `localhost:44355/` (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="b9bb1-165">La forme de l’ID entier doit être `api://localhost:44355/$App ID GUID$`; par exemple`api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-165">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span> 

1. <span data-ttu-id="b9bb1-166">Sélectionnez le bouton **Ajouter une étendue**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-166">Select the **Add a scope** button.</span></span> <span data-ttu-id="b9bb1-167">Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-167">In the panel that opens, enter `access_as_user` as the **Scope name**.</span></span>

1. <span data-ttu-id="b9bb1-168">Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .</span><span class="sxs-lookup"><span data-stu-id="b9bb1-168">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="b9bb1-169">Renseignez les champs pour configurer les invites de consentement des administrateurs et utilisateurs avec les valeurs appropriées pour l’étendue `access_as_user` qui permet à l’application Office hôte d’utiliser l’API web de votre complément avec les mêmes droits que l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-169">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="b9bb1-170">Suggestions :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-170">Suggestions:</span></span>

    - <span data-ttu-id="b9bb1-171">**Titre consentement administrateur** : Office peut agir en tant qu’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-171">**Admin consent title:** Office can act as the user.</span></span>
    - <span data-ttu-id="b9bb1-172">**Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits que l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-172">**Admin consent description:** Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="b9bb1-173">**Titre consentement utilisateur** : Office peut agir à votre place.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-173">**User consent title:** Office can act as you.</span></span>
    - <span data-ttu-id="b9bb1-174">**Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits dont vous disposez.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-174">**Admin consent description:** Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="b9bb1-175">Vérifiez que **State** est défini comme **Activé**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-175">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="b9bb1-176">Sélectionnez **Ajouter une étendue**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-176">Select **Add scope**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b9bb1-177">La partie domaine du **Nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’URI d’ID d’application définie à l’étape précédente avec `/access_as_user`ajouté à la fin, par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-177">The domain part of the Scope name displayed just below the text field should automatically match the Application ID URI set in the previous step, with  appended to the end; for example, .</span></span>

1. <span data-ttu-id="b9bb1-178">Dans la section **Applications client autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-178">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="b9bb1-179">Chacun des ID suivants doit être pré-autorisé.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-179">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="b9bb1-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="b9bb1-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="b9bb1-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office sur le web)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="b9bb1-183">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office sur le web)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-183">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office on the web)</span></span>

    <span data-ttu-id="b9bb1-184">Pour chaque ID, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-184">For each ID, take these steps:</span></span>

    <span data-ttu-id="b9bb1-185">a.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-185">a.</span></span> <span data-ttu-id="b9bb1-186">Sélectionnez le bouton **Ajouter une application client** puis, dans le volet qui s’ouvre, définissez l’ID Client pour le GUID respectif et cochez la case pour `api://localhost:44355/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-186">Select **Add a client application** button then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="b9bb1-187">b.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-187">b.</span></span> <span data-ttu-id="b9bb1-188">Sélectionnez **Ajouter une application**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-188">Select **Add application**.</span></span>

1. <span data-ttu-id="b9bb1-189">Sélectionnez **Autorisations API** sous **Gestion** et sélectionnez **Ajouter une autorisation**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-189">Select **API permissions** under **Manage** and select **Add a permission**.</span></span> <span data-ttu-id="b9bb1-190">Dans le volet qui s’ouvre, sélectionnez **Microsoft Graph**, puis **Autorisations déléguées**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-190">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="b9bb1-191">Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-191">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="b9bb1-192">Sélectionnez les éléments suivants.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-192">Select from the following:</span></span> <span data-ttu-id="b9bb1-193">Votre complément proprement dit ne requiert que la première. Mais l’autorisation `profile` est également requise pour que l’hôte Office puisse obtenir un jeton pour l’application web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-193">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>

    * <span data-ttu-id="b9bb1-194">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="b9bb1-194">Files.Read.All</span></span>
    * <span data-ttu-id="b9bb1-195">profil</span><span class="sxs-lookup"><span data-stu-id="b9bb1-195">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="b9bb1-196">L’autorisation `User.Read` est peut-être déjà répertoriée par défaut.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-196">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="b9bb1-197">Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-197">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="b9bb1-198">Activez la case à cocher pour chaque autorisation telle qu’elle apparaît.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-198">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="b9bb1-199">Après avoir sélectionné les autorisations dont votre complément a besoin, sélectionnez le bouton **Ajouter des autorisations** situé en bas du panneau.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-199">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="b9bb1-200">Sur la même page, sélectionnez le bouton **Accorder l’autorisation d’administrateur pour [nom du client]**, puis **Oui** pour la confirmation qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-200">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.</span></span>

## <a name="configure-the-add-in"></a><span data-ttu-id="b9bb1-201">Configurer le complément</span><span class="sxs-lookup"><span data-stu-id="b9bb1-201">Configure the add-in</span></span>

1. <span data-ttu-id="b9bb1-202">Ouvrez le dossier `\Begin` dans le projet cloné dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-202">Open the `\Begin` folder in the cloned project in your code editor.</span></span>

1. <span data-ttu-id="b9bb1-203">Ouvrez le fichier `.ENV` et utilisez les valeurs que vous avez précédemment copiées.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-203">Open the `.ENV` file and use the values that you copied earlier.</span></span> <span data-ttu-id="b9bb1-204">Configurez la **CLIENT_ID** sur votre **ID d’application (client)** et attribuez la **CLIENT_SECRET** à votre clé secrète client.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-204">Set the **CLIENT_ID** to your **Application (client) ID**, and set the **CLIENT_SECRET** to your client secret.</span></span> <span data-ttu-id="b9bb1-205">Les valeurs ne doivent **pas** se trouver entre des guillemets.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-205">The values should **not** be in quotation marks.</span></span> <span data-ttu-id="b9bb1-206">Quand vous avez terminé, votre modèle doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-206">When you are done, the file should be similar to the following:</span></span> 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. <span data-ttu-id="b9bb1-207">Ouvrez le fichier `\public\javascripts\fallbackAuthDialog.js`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-207">Open the `\public\javascripts\fallbackAuthDialog.js` file.</span></span> <span data-ttu-id="b9bb1-208">Dans la `msalConfig`déclaration, remplacez l’espace réservé $application_GUID here$ par l’ID d’application que vous avez copié lorsque vous avez inscrit votre complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-208">In the `msalConfig` declaration, replace the placeholder $application_GUID here$ with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="b9bb1-209">Les valeurs ne doivent pas être entre guillemets.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-209">The entire notation should be enclosed in quotation marks (").</span></span>

1. <span data-ttu-id="b9bb1-210">Ouvrez le fichier manifeste de complément « manifest\manifest_local. xml », puis faites défiler la page jusqu’à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-210">Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file.</span></span> <span data-ttu-id="b9bb1-211">Juste au-dessus de la `</VersionOverrides>`balise de fin, vous trouverez la marque suivante :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-211">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="b9bb1-212">Remplacez l’espace réservé « $application_GUID here$ » *aux deux endroits* du balisage par l’ID d’application que vous avez copiée lorsque vous avez inscrit votre complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-212">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="b9bb1-213">Les « $ » ne faisant pas partie de l’ID, vous ne devez pas les inclure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-213">The "" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="b9bb1-214">C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-214">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b9bb1-215">La valeur de la **ressource** est l’**URI de l’ID d’application** que vous avez défini lors de l’inscription du complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-215">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span> <span data-ttu-id="b9bb1-216">La section **Étendues** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-216">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="b9bb1-217">Code du côté client</span><span class="sxs-lookup"><span data-stu-id="b9bb1-217">Code the client-side</span></span>

### <a name="create-the-sso-logic"></a><span data-ttu-id="b9bb1-218">Créer la logique SSO</span><span class="sxs-lookup"><span data-stu-id="b9bb1-218">Create the SSO logic</span></span>

1. <span data-ttu-id="b9bb1-219">Ouvrez le fichier `public\javascripts\ssoAuthES6.js` dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-219">In your code editor, open the src\server.ts file.</span></span> <span data-ttu-id="b9bb1-220">Il possède déjà du code qui garantit que les promesses sont prises en charge, même dans Internet Explorer 11, et un appel `Office.onReady` pour attribuer un gestionnaire au bouton unique du complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-220">It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b9bb1-221">Comme leur nom l’indique, ssoAuthES6.js utilise la syntaxe JavaScript ES6, car l’utilisation de `async` et de `await` illustre le mieux la simplicité de l’API SSO.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-221">As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API.</span></span> <span data-ttu-id="b9bb1-222">Lorsque le serveur localhost est démarré, ce fichier est transpilé vers la syntaxe ES5 pour que l’exemple s’exécute dans Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-222">When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will run in Internet Explorer 11.</span></span> 

1. <span data-ttu-id="b9bb1-223">Ajoutez le code suivant sous la méthode Office.onReady :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-223">Add the following code below the Office.onReady method:</span></span>

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exhange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. <span data-ttu-id="b9bb1-224">Remplacez `TODO 1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-224">Replace `TODO 1` with the following code.</span></span> <span data-ttu-id="b9bb1-225">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="b9bb1-225">About this code, note:</span></span>

    - <span data-ttu-id="b9bb1-226">`OfficeRuntime.auth.getAccessToken` commande à Office d’obtenir un jeton de démarrage à partir d’Azure AD.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-226">`OfficeRuntime.auth.getAccessToken` instructs Office to get a bootstrap token from Azure AD.</span></span> <span data-ttu-id="b9bb1-227">Un jeton d’amorçage est semblable à un jeton d’ID, mais il possède une `scp` propriété (étendue) ayant la valeur `access-as-user`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-227">A bootstrap token is similar to an ID token, but it has a `scp` (scope) property with the value `access-as-user`.</span></span> <span data-ttu-id="b9bb1-228">Ce type de jeton peut être échangé par une application Web pour un jeton d’accès à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-228">This kind of token can be exchanged by a web application for an access token to Microsoft Graph.</span></span>
    - <span data-ttu-id="b9bb1-229">Le paramétrage de l’option de `allowSignInPrompt`sur TRUE signifie que si aucun utilisateur n’est actuellement connecté à Office, Office ouvre une invite de connexion contextuelle.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-229">Setting the `allowSignInPrompt`option to true means that if no user is currently signed into Office, then Office will open a popup sign-in prompt.</span></span>
    - <span data-ttu-id="b9bb1-230">Le paramétrage de l’option de `forMSGraphAccess` sur TRUE signale à Office que le complément envisage d’utiliser le jeton de démarrage pour obtenir un jeton d’accès à Microsoft Graph, plutôt que de l’utiliser simplement comme jeton d’ID.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-230">Setting the `forMSGraphAccess` option to true signals to Office that the add-in intends to use the bootstrap token to get an access token to Micrsoft Graph, instead of just using it as an ID token.</span></span> <span data-ttu-id="b9bb1-231">Si l’administrateur du client n’a pas accordé l’autorisation d’accès au complément dans Microsoft Graph, `OfficeRuntime.auth.getAccessToken` renvoie l’erreur **13012**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-231">If the tenant administrator has not granted consent to the add-in's access to Microsoft Graph, then `OfficeRuntime.auth.getAccessToken` returns error **13012**.</span></span> <span data-ttu-id="b9bb1-232">Le complément peut répondre en rétablissant un autre système d’autorisation, ce qui est nécessaire car Office peut uniquement inviter pour accepter le profil Azure AD de l’utilisateur, et non les étendues Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-232">The add-in can respond by falling back to an alternative system of authorization, which is necessary because Office can prompt only for consent to the user's Azure AD profile, not to any Microsoft Graph scopes.</span></span> <span data-ttu-id="b9bb1-233">Le système d’autorisation de secours oblige l’utilisateur à se reconnecter et l’utilisateur *peut* être invité à accepter les étendues de Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-233">The fallback authorization system requires the user to sign in again and the user *can* be prompted to consent to Micrsoft Graph scopes.</span></span> <span data-ttu-id="b9bb1-234">Par conséquent, l’option `forMSGraphAccess` permet de s’assurer que le complément ne fera pas d’échange de jetons échouant en raison d’une absence d’autorisation.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-234">So, the `forMSGraphAccess` option ensures that the add-in won't make a token exchange that will fail due to lack of consent.</span></span> <span data-ttu-id="b9bb1-235">(ayant reçu votre consentement de la part de l’administrateur lors d’une étape précédente, ce scénario ne se produira pas pour ce complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-235">(Since you granted administrator consent in an earlier step, this scenario won't happen for this add-in.</span></span> <span data-ttu-id="b9bb1-236">Mais l’option est tout de même incluse ici pour illustrer les pratiques recommandées.)</span><span class="sxs-lookup"><span data-stu-id="b9bb1-236">But the option is included here anyway to illustrate a best practice.)</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
    ```

1. <span data-ttu-id="b9bb1-237">Remplacez `TODO 2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-237">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="b9bb1-238">Vous créerez la méthode `getGraphToken` lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-238">You'll create the `getGraphToken` method in a later step.</span></span>

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. <span data-ttu-id="b9bb1-239">Remplacez `TODO 3` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-239">Replace `TODO 3` with the following.</span></span> <span data-ttu-id="b9bb1-240">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-240">About this code, note:</span></span> 

    - <span data-ttu-id="b9bb1-241">Si le client Office 365 est configuré pour exiger l’authentification multifacteur, l' `exchangeResponse` inclut une propriété `claims` contenant des informations sur les facteurs supplémentaires requis.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-241">If the Office 365 tenant has been configured to require multifactor authentication, then the `exchangeResponse` will include a `claims` property with information about the additional required factors.</span></span> <span data-ttu-id="b9bb1-242">Dans ce cas, `OfficeRuntime.auth.getAccessToken` doit être rappelé avec l’option `authChallenge` configurée avec la valeur de la propriété revendications.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-242">In that case, `OfficeRuntime.auth.getAccessToken` should be called again with the `authChallenge` option set to the value of the claims property.</span></span> <span data-ttu-id="b9bb1-243">Cela indique à AAD d’inviter l’utilisateur à accepter tous les formulaires d’authentification requis.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-243">This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. <span data-ttu-id="b9bb1-244">Remplacez `TODO 4` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-244">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="b9bb1-245">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-245">About this code, note:</span></span> 

    - <span data-ttu-id="b9bb1-246">Vous créerez la méthode `handleAADErrors` lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-246">You'll create the `handleAADErrors` method in a later step.</span></span> <span data-ttu-id="b9bb1-247">Les erreurs Azure AD sont renvoyées au client sous forme de réponses de code HTTP 200.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-247">Azure AD errors are returned to the client as HTTP code 200 Responses.</span></span> <span data-ttu-id="b9bb1-248">Elles ne génèrent pas d’erreur et ne déclenchent donc pas le `catch`blocage de la`getGraphData` méthode.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-248">They do not throw errors, so they do not trigger the `catch` block of the `getGraphData` method.</span></span>
    - <span data-ttu-id="b9bb1-249">Vous créerez la méthode `makeGraphApiCall` lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-249">You'll create the `makeGraphApiCall` method in a later step.</span></span> <span data-ttu-id="b9bb1-250">Elle effectue un appel AJAX au point de terminaison MS Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-250">It makes an AJAX call to the MS Graph endpoint.</span></span> <span data-ttu-id="b9bb1-251">Les erreurs sont interceptées dans le `.fail` rappel de cet appel, et non dans le bloc `catch` de la méthode `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-251">Errors are caught in the `.fail` callback of that call, not in the `catch` block of the `getGraphData` method.</span></span>

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. <span data-ttu-id="b9bb1-252">Remplacez `TODO 5` par le code suivant</span><span class="sxs-lookup"><span data-stu-id="b9bb1-252">Replace `TODO 5` with the following.</span></span>

    - <span data-ttu-id="b9bb1-253">Les erreurs de l’appel de `getAccessToken` auront une propriété `code` avec un numéro d’erreur généralement dans la plage 13xxx.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-253">Errors from the call of `getAccessToken` will have a `code` property with an error number, typically in the 13xxx range.</span></span> <span data-ttu-id="b9bb1-254">Vous créerez la méthode `handleClientSideErrors` lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-254">You'll create the `handleClientSideErrors` method in a later step.</span></span>
    - <span data-ttu-id="b9bb1-255">La méthode `showMessage` affiche le texte dans le volet Tâches.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-255">The `showMessage` method displays text on the task pane.</span></span>

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. <span data-ttu-id="b9bb1-256">En dessous de la méthode `getGraphData`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-256">Below the `getGraphData` method, add the following.</span></span> <span data-ttu-id="b9bb1-257">Veuillez noter que `/auth` est une route Express côté serveur qui échange le jeton de démarrage avec Azure AD pour un jeton d’accès à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-257">Note that `/auth` is a server-side Express route that exhanges the bootstrap token with Azure AD for an access token to Microsoft Graph.</span></span>

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. <span data-ttu-id="b9bb1-258">En dessous de la méthode `getGraphToken`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-258">Below the `getGraphToken` method, add the following.</span></span> <span data-ttu-id="b9bb1-259">Veuillez noter que `error.code` est un nombre, généralement compris dans la plage 13xxx.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-259">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```
1. <span data-ttu-id="b9bb1-260">Remplacez `TODO 6` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-260">Replace `TODO 6` with the following code.</span></span> <span data-ttu-id="b9bb1-261">Pour plus d’informations sur ces erreurs, reportez-vous à [Résoudre les problèmes liés à SSO dans les compléments Office](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-261">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span> 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the Web.
        showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. <span data-ttu-id="b9bb1-262">Remplacez `TODO 7` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-262">Replace `TODO 7` with the following code.</span></span> <span data-ttu-id="b9bb1-263">Pour plus d’informations sur ces erreurs, reportez-vous à [Résoudre les problèmes liés à SSO dans les compléments Office](troubleshoot-sso-in-office-add-ins.md). La fonction `dialogFallback` appelle le système d’autorisation de secours.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-263">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). The function `dialogFallback` invokes the alternative system of authorization.</span></span> <span data-ttu-id="b9bb1-264">Dans ce complément, le système de secours ouvre une boîte de dialogue demandant à l’utilisateur de se connecter, même si l’utilisateur l’est déjà, et utilise MSAL.js et le flux implicite pour obtenir un jeton d’accès à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-264">In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. <span data-ttu-id="b9bb1-265">Sous la fonction `handleClientSideErrors`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-265">Below the `handleClientSideErrors` function, add the following function.</span></span> 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. <span data-ttu-id="b9bb1-266">Dans de rares cas, le jeton de démarrage qu’Office a mis en cache n’a pas expiré lorsqu’il est validé par Office, mais arrive à expiration au moment où il atteint Azure AD pour l’échange.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-266">On rare occasions the bootstrap token that Office has cached is unexpired when Office validates it, but expires by the time it reaches Azure AD for exchange.</span></span> <span data-ttu-id="b9bb1-267">Azure AD enverra une réponse incluant l’erreur **AADSTS500133**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-267">Azure AD will respond with error **AADSTS500133**.</span></span> <span data-ttu-id="b9bb1-268">Dans ce cas, le complément doit simplement appeler `getGraphData` de manière récursive.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-268">In this case, the add-in should simply recursively call `getGraphData`.</span></span> <span data-ttu-id="b9bb1-269">Le jeton de démarrage mis en cache étant arrivé à expiration, Office en reçoit un nouveau à partir d’Azure AD.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-269">Since the cached bootstrap token is now expired, Office will get a new one from Azure AD.</span></span> <span data-ttu-id="b9bb1-270">Remplacez donc `TODO 8` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-270">So, replace `TODO 8` with the following markup:</span></span> 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)       
    {
        getGraphData();
    }
    ```

1. <span data-ttu-id="b9bb1-271">Pour vous assurer que le complément n’entre pas dans une boucle infinie d’appels vers `getGraphData`, le complément doit effectuer un suivi du nombre de fois où `getGraphData` a été appelé et vérifier qu’il n’est pas appelé de façon récursive plusieurs fois.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-271">To ensure that the add-in doesn't enter an infinite loop of calls to `getGraphData`, the add-in should keep track of how many times `getGraphData` has been called and be sure that is not called recursively called more than once.</span></span> <span data-ttu-id="b9bb1-272">Par conséquent, créez une variable de compteur dans une étendue globale aux fonctions de `handleAADErrors` et `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-272">So, create a counter variable in a scope that is global to the `handleAADErrors` and `getGraphData` functions.</span></span> <span data-ttu-id="b9bb1-273">Un bon emplacement pour les variables globales se trouve juste en dessous de l’appel de méthode `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-273">A good place for global variables is just below the `Office.onReady` method call.</span></span>

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. <span data-ttu-id="b9bb1-274">Modifiez la structure `if` dans la méthode `handleAADErrors` de façon à ce qu’elle :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-274">Change the `if` structure in the `handleAADErrors` method so that it:</span></span>

    - <span data-ttu-id="b9bb1-275">Incrémente le compteur juste avant qu’il n’appelle `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-275">Increments the counter just before it calls `getGraphData`.</span></span>
    - <span data-ttu-id="b9bb1-276">Vérifie que `getGraphData` n’a pas déjà été appelé une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-276">Tests to ensure that `getGraphData` has not already been called a second time.</span></span> 

    <span data-ttu-id="b9bb1-277">La version finale de la structure `if` doit donc ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-277">So the final version of the `if` structure should look like the following:</span></span>

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="b9bb1-278">Remplacez `TODO 9` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-278">Replace `TODO 9` with the following.</span></span> 

    ```javascript
    else {                
        dialogFallback();
    }
    ```

1. <span data-ttu-id="b9bb1-279">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-279">Save and close the file.</span></span>

### <a name="get-the-data-and-add-it-to-the-office-document"></a><span data-ttu-id="b9bb1-280">Obtenir les données et les ajouter au document Office</span><span class="sxs-lookup"><span data-stu-id="b9bb1-280">Get the data and add it to the Office document</span></span>

1. <span data-ttu-id="b9bb1-281">Dans le dossier `public\javascripts`, créez un fichier appelé `data.js`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-281">In the `public\javascripts` folder, create a new file named `data.js`, and paste the following code:</span></span>

1. <span data-ttu-id="b9bb1-282">Ajoutez la fonction suivante au fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-282">Add the following function to the file.</span></span> <span data-ttu-id="b9bb1-283">Il s’agit de la fonction appelée par la fonction `getGraphData` lorsqu’elle a acquis un jeton d’accès pour Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-283">This is the function that is called by the `getGraphData` function when it has acquired an access token to Microsoft Graph.</span></span> 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. <span data-ttu-id="b9bb1-284">Remplacez `TODO 10` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-284">Replace `TODO 10` with the following.</span></span> <span data-ttu-id="b9bb1-285">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-285">About this code, note:</span></span> 

    - <span data-ttu-id="b9bb1-286">Cet objet est le paramètre de la méthode `$.ajax`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-286">This object is the parameter to the `$.ajax` method.</span></span>
    - <span data-ttu-id="b9bb1-287">Le `/getuserdata` est une route Express sur le serveur du complément que vous créez au cours d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-287">The `/getuserdata` is an Express route on the add-in's server that you create in a later step.</span></span> <span data-ttu-id="b9bb1-288">Elle appellera un point de terminaison Microsoft Graph et inclura le jeton d’accès dans son appel.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-288">It will call a Microsoft Graph endpoint and include the access token in its call.</span></span> 

    ```javascript
    {
        type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. <span data-ttu-id="b9bb1-289">Remplacez `TODO11` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-289">Replace `TODO11` with the following.</span></span> <span data-ttu-id="b9bb1-290">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-290">About this code, note:</span></span>

    - <span data-ttu-id="b9bb1-291">Le `writeFileNamesToOfficeDocument` insère les données de Graph dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-291">The `writeFileNamesToOfficeDocument` will insert the data from Graph into the Office document.</span></span> <span data-ttu-id="b9bb1-292">Il est défini dans le fichier `public\javascripts\document.js`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-292">The `public\javascripts\document.js` method is defined in the src\auth.ts file.</span></span> 
    - <span data-ttu-id="b9bb1-293">Si `writeFileNamesToOfficeDocument` renvoie une erreur, il commence par « Impossible d’ajouter des noms de fichiers au document ».</span><span class="sxs-lookup"><span data-stu-id="b9bb1-293">If `writeFileNamesToOfficeDocument` returns an error, it will begin with "Unable to add filenames to document."</span></span>

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {        
        showMessage(error);
    });
    ```

1. <span data-ttu-id="b9bb1-294">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-294">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="b9bb1-295">Code du côté serveur</span><span class="sxs-lookup"><span data-stu-id="b9bb1-295">Code the server-side</span></span>

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a><span data-ttu-id="b9bb1-296">Créer le routeur d’authentification et la logique d’échange de jetons</span><span class="sxs-lookup"><span data-stu-id="b9bb1-296">Create the auth router and the token exchange logic</span></span>

1. <span data-ttu-id="b9bb1-297">Ouvrez le fichier `routes\authRoute.js` et ajoutez la fonction d’itinéraire suivante juste en dessous des instructions `require` et au-dessus de l’instruction `module.exports`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-297">Open the file `routes\authRoute.js` and add the following route function just below the `require` statements and above the `module.exports` statement.</span></span> <span data-ttu-id="b9bb1-298">Veuillez noter que le paramètre d’URL de `router.get` est'/'.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-298">Note that the URL parameter of `router.get` is '/'.</span></span> <span data-ttu-id="b9bb1-299">Cet itinéraire étant défini dans un routeur qui gère toutes les requêtes HTTP pour l’URL « /auth », il gère toutes les demandes pour « /auth ».</span><span class="sxs-lookup"><span data-stu-id="b9bb1-299">Since this route is being defined in a router that will handle all HTTP Requests for the URL '/auth', this route effectively handles all requests for '/auth'.</span></span> <span data-ttu-id="b9bb1-300">La fonction `getGraphToken` côté client créée précédemment appelle cet itinéraire.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-300">The client-side `getGraphToken` function that you created earlier calls this route.</span></span>  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exhange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. <span data-ttu-id="b9bb1-301">Remplacez `TODO 12` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-301">Replace `TODO 12` with the following code.</span></span>

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. <span data-ttu-id="b9bb1-302">Remplacez `TODO 13` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-302">Replace `TODO 13` with the following code.</span></span> <span data-ttu-id="b9bb1-303">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="b9bb1-303">About this code, note:</span></span> 

    - <span data-ttu-id="b9bb1-304">Il s’agit du début d’un long `else` bloc, mais la fermeture `}` n’est pas encore terminée car vous y ajouterez d’autres codes.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-304">This is the beginning of a long `else` block, but the closing `}` is not at the end yet because you will be adding more code to it.</span></span> 
    - <span data-ttu-id="b9bb1-305">La chaîne de `authorization` est « Porteur » suivi du jeton de démarrage, de sorte que la première ligne du bloc `else` attribue le jeton à la `jwt`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-305">The `authorization` string is "Bearer " followed by the bootstrap token, so the first line of the `else` block is assigning the token to the `jwt`.</span></span> <span data-ttu-id="b9bb1-306">(« JWT » signifie « jeton Web JSON »).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-306">("JWT" stands for "JSON Web Token".)</span></span>
    - <span data-ttu-id="b9bb1-307">Les deux valeurs `process.env.*` sont les constantes que vous avez attribuées lors de la configuration du complément.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-307">The two `process.env.*` values are the constants that you assigned when you configured the add-in.</span></span> 
    - <span data-ttu-id="b9bb1-308">Le paramètre de formulaire `requested_token_use` est paramétré sur « On_behalt_of ».</span><span class="sxs-lookup"><span data-stu-id="b9bb1-308">The `requested_token_use` form parameter is set to 'on_behalf_of'.</span></span> <span data-ttu-id="b9bb1-309">Cette option indique à Azure AD que le complément demande un jeton d’accès à Microsoft Graph à l’aide du flux On-Behalf-Of.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-309">This tells Azure AD that the add-in is requesting an access token to Microsoft Graph using the On-Behalf-Of Flow.</span></span> <span data-ttu-id="b9bb1-310">Azure répond en validant que le jeton de démarrage, affecté au paramètre de formulaire `assertion`, a une propriété `scp` configurée sur `access-as-user`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-310">Azure will respond by validating that the bootstrap token, which is assigned to `assertion` form parameter, has a `scp` property that is set to `access-as-user`.</span></span>
    - <span data-ttu-id="b9bb1-311">Le paramètre de formulaire `scope` est défini sur « Files.Read.All », qui est la seule étendue Microsoft Graph dont le complément a besoin.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-311">The `scope` form parameter is set to 'Files.Read.All' which is the only Microsoft Graph scope that the add-in needs.</span></span>

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. <span data-ttu-id="b9bb1-312">Remplacez `TODO 14` par le code suivant, qui termine le bloc `else`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-312">Replace `TODO 14` with the following code, which completes the `else` block.</span></span> <span data-ttu-id="b9bb1-313">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-313">About this code, note:</span></span>

    - <span data-ttu-id="b9bb1-314">Le `tenant` const est défini sur « commun », car vous avez configuré le complément en tant que multiclient lorsque vous l’avez inscrit avec Azure AD, en particulier lorsque vous configurez **types de compte pris en charge** pour **les comptes de n’importe quel annuaire d’organisation et les comptes Microsoft personnels (par exemple, Skype, Xbox, Outlook.com)**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-314">The const `tenant` is set to 'common' because you configured the add-in as multitenant when you registered it with Azure AD; specifically when you set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span> <span data-ttu-id="b9bb1-315">Si vous avez en revanche choisi de prendre en charge uniquement les comptes figurant dans la même location Office 365 que le complément enregistré, `tenant` dans ce code serait défini sur le GUID du client.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-315">If you had instead chosen to support only accounts in the same Office 365 tenancy where the add-in is registered, then in this code `tenant` would be set to the GUID of the tenant.</span></span> 
    - <span data-ttu-id="b9bb1-316">Si la requête POST ne génère pas d’erreur, la réponse d’Azure AD est convertie en JSON et envoyée au client.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-316">If the POST request does not error, then the response from Azure AD is converted to JSON and sent to the client.</span></span> <span data-ttu-id="b9bb1-317">Cet objet JSON possède une propriété `access_token` à laquelle Azure AD a attribué un jeton d’accès à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-317">This JSON object has an `access_token` property to which Azure AD has assigned the access token to Microsoft Graph.</span></span>

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();
            
            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. <span data-ttu-id="b9bb1-318">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-318">Save and close the file.</span></span>

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a><span data-ttu-id="b9bb1-319">Créer l’itinéraire qui permettra de récupérer les données à partir de Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="b9bb1-319">Create the route that will fetch the data from Microsoft Graph</span></span>

1. <span data-ttu-id="b9bb1-320">Ouvrez le fichier `app.js` dans la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-320">Open the Startup.cs file in the root of the project.</span></span> <span data-ttu-id="b9bb1-321">Juste en dessous de la route pour « /Dialog.html », ajoutez l’itinéraire suivant.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-321">Just below the route for '/dialog.html', add the following route.</span></span> <span data-ttu-id="b9bb1-322">Cet itinéraire est appelé par la fonction `makeGraphApiCall` que vous avez créée lors d’une étape précédente.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-322">This route is called by the `makeGraphApiCall` function that you created in an earlier step.</span></span>

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. <span data-ttu-id="b9bb1-323">Remplacez `TODO 15` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-323">Replace `TODO 15` with the following.</span></span> <span data-ttu-id="b9bb1-324">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-324">About this code, note:</span></span>

    - <span data-ttu-id="b9bb1-325">L’appelant de cet itinéraire, `makeGraphApiCall`, a ajouté un jeton d’accès à Microsoft Graph à la demande HTTP en tant qu’en-tête nommé « access_token ».</span><span class="sxs-lookup"><span data-stu-id="b9bb1-325">The caller of this route, `makeGraphApiCall`, added the access token to Microsoft Graph to the HTTP Request as a header named "access_token".</span></span>
    - <span data-ttu-id="b9bb1-326">La fonction de `getGraphData` est défini dans le fichier `msgraph-helper.js`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-326">The  method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="b9bb1-327">(il ne s’agit pas de la même fonction que la fonction `getGraphData` côté client que vous avez définie dans le fichier `ssoAuthES6.js`).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-327">(This is not the same function as the client-side `getGraphData` function that you defined in the `ssoAuthES6.js` file.)</span></span>
    - <span data-ttu-id="b9bb1-328">Le dernier paramètre pour `queryParamsSegment` est codé en dur.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-328">The last parameter, for `queryParamsSegment`, is hardcoded.</span></span> <span data-ttu-id="b9bb1-329">Si vous modifiez ce code dans un complément production et qu’une partie quelconque de `queryParamsSegment` provient d’une intervention de l’utilisateur, n’oubliez pas qu’il est purgé afin qu’il ne puisse pas être utilisé dans une attaque par injection d’en-tête de réponse.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-329">If you reuse this code in a production add-in and any part of `queryParamsSegment` comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.</span></span>
    - <span data-ttu-id="b9bb1-330">Le code minimise les données qui doivent provenir de Microsoft Graph en spécifiant uniquement la propriété nécessaire (« nom ») et uniquement les 10 premiers noms de dossier ou de fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-330">The code minimizes the data that must come from Microsoft Graph by specifying only the property we need ("name") and only the top 10 folder or file names.</span></span>

    ```javascript
    const graphToken = req.get('access_token');    
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. <span data-ttu-id="b9bb1-331">Remplacez `TODO 16` par ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-331">Replace `TODO 16` with the following.</span></span> <span data-ttu-id="b9bb1-332">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b9bb1-332">About this code, note:</span></span>

    - <span data-ttu-id="b9bb1-333">Si Microsoft Graph renvoie une erreur, un jeton non valide ou expiré par exemple, une propriété de code dans l’objet renvoyé est attribuée à un état HTTP (par exemple, 401).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-333">If Microsoft Graph returns an error, such as invalid or expired token, there will be a code property in the returned object set to a HTTP status (e.g., 401).</span></span> <span data-ttu-id="b9bb1-334">Le code relaie l’erreur vers le client.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-334">The code relays the error to the client.</span></span> <span data-ttu-id="b9bb1-335">Elle sera interceptée dans le `.fail` rappel de `makeGraphApiCall`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-335">It will be caught in the `.fail` callback of `makeGraphApiCall`.</span></span>
    - <span data-ttu-id="b9bb1-336">Les données Microsoft Graph incluent des métadonnées OData et des eTags dont le complément n’a pas besoin, de sorte que le code construit un nouveau groupe contenant uniquement le noms des fichiers à envoyer au client.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-336">Microsoft Graph data includes OData metadata and eTags that the add-in does not need, so the code constructs a new array containing only the file names to send to the client.</span></span>

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. <span data-ttu-id="b9bb1-337">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-337">Save and close the file.</span></span>

## <a name="run-the-project"></a><span data-ttu-id="b9bb1-338">Exécutez le projet</span><span class="sxs-lookup"><span data-stu-id="b9bb1-338">Run the project</span></span>

1. <span data-ttu-id="b9bb1-339">Assurez-vous d’avoir des fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-339">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="b9bb1-340">Ouvrez une invite de commandes dans la racine du dossier `\Complete`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-340">Open a command prompt in the root of the `\Complete` folder.</span></span> 

1. <span data-ttu-id="b9bb1-341">Exécutez la commande `npm start`.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-341">Run the command  `npm start`.</span></span> 

1. <span data-ttu-id="b9bb1-342">Vous devez charger une version du complément dans une application Office (Excel, Word ou PowerPoint) pour le tester.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-342">You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it.</span></span> <span data-ttu-id="b9bb1-343">Les instructions sont fonction de votre plateforme.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-343">The instructions depend on your platform.</span></span> <span data-ttu-id="b9bb1-344">Vous trouverez des liens vers des instructions sur [Charger une version du complément Office pour le tester](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span><span class="sxs-lookup"><span data-stu-id="b9bb1-344">There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span></span>

1. <span data-ttu-id="b9bb1-345">Dans l’application Office, sur le ruban **Accueil**, sélectionnez le bouton **Afficher le complément** dans le groupe **Node.js SSO** pour ouvrir le complément du panneau des tâches.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-345">In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.</span></span>

1. <span data-ttu-id="b9bb1-346">Cliquez sur le bouton **Obtenir des noms de fichier OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-346">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="b9bb1-347">Si vous êtes connecté à Office à l’aide d’un compte professionnel ou scolaire (Office 365) ou d’un compte Microsoft et que l’authentification unique fonctionne comme prévu, les 10 premiers noms de fichier et de dossiers dans votre espace OneDrive Entreprise sont insérés dans le document.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-347">If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document.</span></span> <span data-ttu-id="b9bb1-348">(la première opération peut prendre jusqu’à 15 secondes). Si vous n’êtes pas connecté ou si vous êtes dans un scénario qui ne prend pas en charge SSO ou si l’authentification unique ne fonctionne pas pour une raison quelconque, vous serez invité à vous connecter.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-348">(It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="b9bb1-349">Une fois connecté, les noms de fichier et de dossier s’affichent.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-349">After you log in, the file and folder names appear.</span></span>

> [!NOTE]
> <span data-ttu-id="b9bb1-350">Si vous étiez précédemment connecté à Office avec un ID différent et si certaines applications précédemment ouvertes Office le sont toujours, Office ne changera pas systématiquement votre identifiant même si cela semble être le cas.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-350">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="b9bb1-351">Dans ce cas, l’appel vers Microsoft Graph peut échouer ou des données de l’ID précédent peuvent être renvoyées.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-351">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="b9bb1-352">Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir des noms de fichiers OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="b9bb1-352">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
