---
title: Créer un complément Office ASP.NET qui utilise l’authentification unique
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: bc8c2427171f06865de6c809a5d7311018fcc278
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695804"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="5d388-102">Créer un complément Office ASP.NET qui utilise l’authentification unique (aperçu)</span><span class="sxs-lookup"><span data-stu-id="5d388-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="5d388-p101">Lorsque les utilisateurs sont connectés à Office, votre complément peut utiliser les mêmes informations d’identification pour permettre aux utilisateurs d’accéder à plusieurs applications sans avoir à se connecter une deuxième fois. Pour en savoir plus, consultez [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="5d388-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="5d388-105">Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré avec ASP.NET, OWIN et la bibliothèque d’authentification Microsoft (MSAL) pour .NET.</span><span class="sxs-lookup"><span data-stu-id="5d388-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="5d388-106">Pour un article similaire concernant un complément basé sur Node.js, consultez [Création d’un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="5d388-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5d388-107">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="5d388-107">Prerequisites</span></span>

* <span data-ttu-id="5d388-108">Version la plus récente disponible de Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="5d388-108">The latest available version of Visual Studio 2017.</span></span>

* <span data-ttu-id="5d388-109">Office 365 (version d’Office par abonnement).</span><span class="sxs-lookup"><span data-stu-id="5d388-109">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="5d388-110">Dernière version mensuelle et build du canal du programme Insider.</span><span class="sxs-lookup"><span data-stu-id="5d388-110">Latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="5d388-111">Vous devez participer au programme Office Insider pour obtenir cette version.</span><span class="sxs-lookup"><span data-stu-id="5d388-111">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="5d388-112">Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="5d388-112">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="5d388-113">Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.</span><span class="sxs-lookup"><span data-stu-id="5d388-113">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="5d388-114">Configurer le projet de démarrage</span><span class="sxs-lookup"><span data-stu-id="5d388-114">Set up the starter project</span></span>

1. <span data-ttu-id="5d388-115">Clonez ou téléchargez le référentiel sur [Complément Office ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="5d388-115">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="5d388-p103">Ouvrez le dossier **Before** et ouvrez le fichier .sln dans Visual Studio. Il s’agit d’un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés.</span><span class="sxs-lookup"><span data-stu-id="5d388-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5d388-p104">Il existe également une version finale de l’échantillon dans le même référentiel. Elle est équivalente au complément que vous obtiendriez si vous terminiez les procédures de cet article, sauf que le projet terminé comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, ouvrez simplement le fichier `sln` et suivez les instructions de cet article, mais ignorez les sections **Code côté client** et **Code côté serveur**.</span><span class="sxs-lookup"><span data-stu-id="5d388-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="5d388-p105">Une fois le projet ouvert, générez-le dans Visual Studio, qui installera les packages répertoriés dans le fichier packages.config. L’opération peut prendre de quelques secondes à plusieurs minutes selon le nombre de packages présents dans le cache de packages de l’ordinateur local.</span><span class="sxs-lookup"><span data-stu-id="5d388-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5d388-p106">Vous obtiendrez une erreur relative à l’espace de noms Identity. Il s’agit d’un effet indésirable dû à un problème de configuration qui sera corrigé à la prochaine étape. Le plus important est que les packages soient bien installés.</span><span class="sxs-lookup"><span data-stu-id="5d388-p106">You will get an error about the Identity namespace. This is a side effect of a configuration issue that you will fix with the next step. The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="5d388-127">Pour l’instant, la version de la bibliothèque MSAL (Microsoft.Identity.Client) dont vous avez besoin pour l’authentification unique (version `1.1.4-preview0002`) ne fait pas partie du catalogue NuGet standard, elle n’est donc pas répertoriée dans package.config et doit être installée séparément.</span><span class="sxs-lookup"><span data-stu-id="5d388-127">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.4-preview0002`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span>

   > 1. <span data-ttu-id="5d388-128">Dans le menu **Outils**, accédez à **Gestionnaire de package NuGet** > **Console du Gestionnaire de package**.</span><span class="sxs-lookup"><span data-stu-id="5d388-128">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span>
   > 2. <span data-ttu-id="5d388-129">Dans la console, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="5d388-129">At the console, run the following command.</span></span> <span data-ttu-id="5d388-130">L’opération peut prendre une minute ou plus, même avec une bonne connexion Internet.</span><span class="sxs-lookup"><span data-stu-id="5d388-130">It may take a minute or more to complete even with a fast Internet connection.</span></span> <span data-ttu-id="5d388-131">Une fois l’opération terminée, le message **Successfully installed ’Microsoft.Identity.Client 1.1.4-preview0002’ ...** doit être affiché vers la fin de la sortie de la console.</span><span class="sxs-lookup"><span data-stu-id="5d388-131">When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.4-preview0002' ...** near the end of the output in the console.</span></span>
   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`
   > 3. <span data-ttu-id="5d388-132">Dans l’**Explorateur de solutions**, développez les **Références** du projet **Office-Add-in-ASPNET-SSO-WebAPI**.</span><span class="sxs-lookup"><span data-stu-id="5d388-132">In **Solution Explorer**, expand **References** of **Office-Add-in-ASPNET-SSO-WebAPI** project.</span></span> <span data-ttu-id="5d388-133">Vérifiez que **Microsoft.Identity.Client** est répertorié.</span><span class="sxs-lookup"><span data-stu-id="5d388-133">Verify that **Microsoft.Identity.Client** is listed.</span></span> <span data-ttu-id="5d388-134">S’il n’y est pas ou qu’une icône d’avertissement figure sur son entrée, supprimez l’entrée, puis utilisez l’Assistant Ajouter une référence Visual Studio pour ajouter une référence à l’assembly dans **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span><span class="sxs-lookup"><span data-stu-id="5d388-134">If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference Wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="5d388-135">Créez le projet une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="5d388-135">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="5d388-136">Enregistrez le complément avec le point de terminaison Azure AD v2.0</span><span class="sxs-lookup"><span data-stu-id="5d388-136">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="5d388-137">Les instructions suivantes présentant un manière générique, vous pouvez les utiliser dans plusieurs emplacements.</span><span class="sxs-lookup"><span data-stu-id="5d388-137">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="5d388-138">En lien avec ce article, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="5d388-138">For this article do the following:</span></span>

- <span data-ttu-id="5d388-139">Remplacez l’espace réservé **$ADD-IN-NAME$** par `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="5d388-139">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="5d388-140">Remplacez l’espace réservé **$FQDN-WITHOUT-PROTOCOL$** par `localhost:44355`.</span><span class="sxs-lookup"><span data-stu-id="5d388-140">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="5d388-141">Lorsque vous spécifiez des autorisations dans la boîte de dialogue **Sélectionner les autorisations**, cochez les cases correspondant aux autorisations suivantes.</span><span class="sxs-lookup"><span data-stu-id="5d388-141">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="5d388-142">Seule la première est réellement nécessaire pour votre complément proprement dit, mais la bibliothèque MSAL utilisée par le code côté serveur requiert `offline_access` et `openid`.</span><span class="sxs-lookup"><span data-stu-id="5d388-142">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="5d388-143">L’autorisation `profile` est requise pour l’hôte Office afin d’obtenir un jeton pour l’application web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-143">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
  * <span data-ttu-id="5d388-144">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="5d388-144">Files.Read.All</span></span>
  * <span data-ttu-id="5d388-145">offline_access</span><span class="sxs-lookup"><span data-stu-id="5d388-145">offline_access</span></span>
  * <span data-ttu-id="5d388-146">openid</span><span class="sxs-lookup"><span data-stu-id="5d388-146">openid</span></span>
  * <span data-ttu-id="5d388-147">profil</span><span class="sxs-lookup"><span data-stu-id="5d388-147">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="5d388-148">Octroi du consentement administrateur pour le complément</span><span class="sxs-lookup"><span data-stu-id="5d388-148">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="5d388-149">Configurer le complément</span><span class="sxs-lookup"><span data-stu-id="5d388-149">Configure the add-in</span></span>

1. <span data-ttu-id="5d388-150">Dans la chaîne suivante, remplacez l’espace réservé “{tenant_ID}” par votre ID de client Office 365.</span><span class="sxs-lookup"><span data-stu-id="5d388-150">In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenancy ID.</span></span> <span data-ttu-id="5d388-151">Si vous n’avez pas copié l’ID de client lorsque vous avez inscrit le complément auprès d’AAD, utilisez une des méthodes dans [Trouver votre ID de client Office 365](/onedrive/find-your-office-365-tenant-id) pour l’obtenir.</span><span class="sxs-lookup"><span data-stu-id="5d388-151">If you didn't copy the tenancy ID when you registered the add-in with AAD, use one of the methods in [Find your Office 365 tenant ID](/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. <span data-ttu-id="5d388-152">Dans Visual Studio, ouvrez le fichier web.config. Il existe certaines clés dans la section **appSettings** à laquelle vous devez affecter des valeurs.</span><span class="sxs-lookup"><span data-stu-id="5d388-152">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

1. <span data-ttu-id="5d388-p112">Utilisez la chaîne que vous avez créée à l’étape 1 en tant que valeur pour la clé nommée « ida:Issuer ». Assurez-vous que la valeur ne comporte aucun espace vide.</span><span class="sxs-lookup"><span data-stu-id="5d388-p112">Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.</span></span>

1. <span data-ttu-id="5d388-155">Affectez les valeurs suivantes aux clés correspondantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-155">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="5d388-156">Clé</span><span class="sxs-lookup"><span data-stu-id="5d388-156">Key</span></span>|<span data-ttu-id="5d388-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d388-157">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="5d388-158">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="5d388-158">ida:ClientID</span></span>|<span data-ttu-id="5d388-159">L’ID d’application que vous avez obtenu lorsque vous avez enregistré le complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="5d388-160">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="5d388-160">ida:Audience</span></span>|<span data-ttu-id="5d388-161">L’ID d’application que vous avez obtenu lorsque vous avez enregistré le complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-161">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="5d388-162">ida:Password</span><span class="sxs-lookup"><span data-stu-id="5d388-162">ida:Password</span></span>|<span data-ttu-id="5d388-163">Mot de passe que vous avez obtenu lorsque vous avez inscrit le complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-163">The password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="5d388-p113">Voici un exemple de ce à quoi doivent ressembler les quatre clés que vous avez modifiées. *Vous remarquerez que les clés ClientID et Audience sont identiques*. Vous pouvez également utiliser une seule clé pour les deux fonctions, mais votre balisage web.config sera mieux réutilisable si vous les séparez, car elles ne sont pas toujours identiques. En outre, des clés séparées renforcent l’idée que votre complément est à la fois une ressource OAuth, par rapport à l’hôte Office, et un client OAuth, par rapport à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="5d388-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />

    ```

   > [!NOTE]
   > <span data-ttu-id="5d388-168">Conservez tels quels les autres paramètres de la section **appSettings**.</span><span class="sxs-lookup"><span data-stu-id="5d388-168">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="5d388-169">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-169">Save and close the file.</span></span>

1. <span data-ttu-id="5d388-170">Dans le projet de complément, ouvrez le fichier manifeste du complément « Office-Add-in-ASPNET-SSO.xml ».</span><span class="sxs-lookup"><span data-stu-id="5d388-170">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span>

1. <span data-ttu-id="5d388-171">Faites défiler vers le bas du fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-171">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="5d388-172">Juste au-dessus de la balise de fin `</VersionOverrides>`, vous trouverez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="5d388-172">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="5d388-p114">Remplacez l’espace réservé « {application_GUID here} » *aux deux endroits* du balisage par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément. Les crochets « {} » ne font pas partie de l’ID, ne les incluez pas. C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.</span><span class="sxs-lookup"><span data-stu-id="5d388-p114">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in. The "{}" are not part of the ID, so do not include them. This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="5d388-176">La valeur **Resource** correspond à l’**URI d’ID d’application** défini lorsque vous avez ajouté la plateforme d’API web à l’enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-176">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="5d388-177">La section **Scopes** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.</span><span class="sxs-lookup"><span data-stu-id="5d388-177">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="5d388-178">Ouvrez l’onglet **Avertissements** de la **liste d’erreurs** dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5d388-178">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="5d388-179">Si un message d’avertissement indique que `<WebApplicationInfo>` n’est pas un enfant valide de `<VersionOverrides>`, votre version de Visual Studio 2017 Preview ne reconnaît pas le balisage d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="5d388-179">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not recognize the SSO markup.</span></span> <span data-ttu-id="5d388-180">Solution de contournement : procédez comme suit pour un complément Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="5d388-180">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="5d388-181">(Si vous utilisez un complément Outlook, consultez la solution de contournement ci-dessous.)</span><span class="sxs-lookup"><span data-stu-id="5d388-181">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="5d388-182">**Solution de contournement pour Word, Excel et PowerPoint**</span><span class="sxs-lookup"><span data-stu-id="5d388-182">**Workaround for Word, Excel, and PowerPoint**</span></span>

        1. <span data-ttu-id="5d388-183">Commentez la section `<WebApplicationInfo>` du manifeste juste au-dessus de la fin de `</VersionOverrides>`.</span><span class="sxs-lookup"><span data-stu-id="5d388-183">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="5d388-p116">Appuyez sur **F5** pour démarrer une session de débogage. Cette opération entraîne la création d’une copie du manifeste dans le dossier suivant (auquel il est plus facile d’accéder dans l’**Explorateur de fichiers** que dans Visual Studio) : `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="5d388-p116">Press **F5** to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="5d388-186">Dans la copie du manifeste, supprimez la syntaxe de commentaire autour de la section `<WebApplicationInfo>`.</span><span class="sxs-lookup"><span data-stu-id="5d388-186">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="5d388-187">Enregistrez la copie du manifeste.</span><span class="sxs-lookup"><span data-stu-id="5d388-187">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="5d388-p117">À présent, vous devez empêcher Visual Studio de remplacer la copie du manifeste la prochaine fois que vous appuyez sur F5. Cliquez avec le bouton droit de la souris sur le nœud de solution en haut de l’**explorateur de solutions** (et non sur l’un des nœuds de projet).</span><span class="sxs-lookup"><span data-stu-id="5d388-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="5d388-190">Sélectionnez **Propriétés** dans le menu contextuel, puis une boîte de dialogue **Pages de propriétés de la solution** s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="5d388-190">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="5d388-191">Développez **Propriétés de configuration** et sélectionnez **Configuration**.</span><span class="sxs-lookup"><span data-stu-id="5d388-191">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="5d388-192">Désélectionnez **Créer** et **Déployer** dans la ligne pour le projet **Office-Add-in-ASPNET-SSO** (et *pas* le projet **Office-Add-in-ASPNET-SSO-WebAPI**).</span><span class="sxs-lookup"><span data-stu-id="5d388-192">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="5d388-193">Cliquez sur **OK** pour fermer la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="5d388-193">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="5d388-194">**Solution de contournement pour Outlook**</span><span class="sxs-lookup"><span data-stu-id="5d388-194">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="5d388-p118">Sur votre ordinateur de développement, recherchez l’élément `MailAppVersionOverridesV1_1.xsd` existant. Il doit se trouver dans le répertoire d’installation Visual Studio sous `./Xml/Schemas/{lcid}`. Par exemple, sur une installation standard de VS 2017 32 bits sur un système anglais (États-Unis), le chemin d’accès complet serait `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span><span class="sxs-lookup"><span data-stu-id="5d388-p118">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="5d388-198">Renommez le fichier existant comme suit : `MailAppVersionOverridesV1_1.old`.</span><span class="sxs-lookup"><span data-stu-id="5d388-198">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="5d388-199">Copiez la version modifiée du fichier dans le dossier : [Schéma MailAppVersionOverrides modifié](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="5d388-199">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="5d388-200">Enregistrez et fermez le fichier manifeste principal dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5d388-200">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="5d388-201">Code côté client</span><span class="sxs-lookup"><span data-stu-id="5d388-201">Code the client side</span></span>

1. <span data-ttu-id="5d388-p119">Ouvrez le fichier Home.js dans le dossier **Scripts**. Il contient déjà du code :</span><span class="sxs-lookup"><span data-stu-id="5d388-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="5d388-204">Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="5d388-204">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="5d388-205">Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.</span><span class="sxs-lookup"><span data-stu-id="5d388-205">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="5d388-206">Une méthode `logErrors` qui consigne dans la console les erreurs qui ne sont pas destinées à l’utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="5d388-206">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="5d388-p120">En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="5d388-p121">La gestion des erreurs dans le complément tente parfois automatiquement d’obtenir un jeton d’accès une deuxième fois, à l’aide d’un autre jeu d’options. La variable de compteur `timesGetOneDriveFilesHasRun` et la variable d’indicateur `triedWithoutForceConsent` permettent de s’assurer que l’utilisateur ne tente pas de manière répétée d’obtenir un jeton sans y parvenir.</span><span class="sxs-lookup"><span data-stu-id="5d388-p121">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="5d388-p122">Vous allez créer la méthode `getDataWithToken` à l’étape suivante, mais rappelez-vous qu’elle définit une option appelée `forceConsent` sur `false`. Vous en saurez plus à la prochaine étape.</span><span class="sxs-lookup"><span data-stu-id="5d388-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```js
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. <span data-ttu-id="5d388-p123">En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="5d388-215">[getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) est la nouvelle API d’Office.js qui permet à un complément de demander à l’application hôte Office (Excel, PowerPoint, Word, etc.) un jeton d’accès au complément (pour l’utilisateur connecté à Office).</span><span class="sxs-lookup"><span data-stu-id="5d388-215">The [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="5d388-216">L’application hôte Office demande alors le jeton au point de terminaison Azure AD 2.0.</span><span class="sxs-lookup"><span data-stu-id="5d388-216">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="5d388-217">Dans la mesure où vous avez préalablement autorisé l’hôte Office sur votre complément lors de son inscription, Azure AD enverra le jeton.</span><span class="sxs-lookup"><span data-stu-id="5d388-217">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="5d388-218">Si aucun utilisateur n’est connecté à Office, l’hôte Office invite l’utilisateur à se connecter.</span><span class="sxs-lookup"><span data-stu-id="5d388-218">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="5d388-p125">Le paramètre d’options définit `forceConsent` sur `false`, donc l’utilisateur ne sera pas invité à accorder à l’hôte Office l’accès à votre complément chaque fois qu’il utilisera le complément. La première fois que l’utilisateur exécutera le complément, l’appel à `getAccessTokenAsync` échouera, mais la logique de gestion des erreurs que vous ajouterez dans une étape ultérieure effectuera automatiquement un autre appel avec le jeu d’options `forceConsent` défini sur `true`, et l’utilisateur sera invité à donner son consentement, mais uniquement la première fois.</span><span class="sxs-lookup"><span data-stu-id="5d388-p125">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in. The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="5d388-221">Vous créerez la méthode `handleClientSideErrors` à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="5d388-221">You will create the `handleClientSideErrors` method in a later step.</span></span>

    ```js
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. <span data-ttu-id="5d388-p126">Remplacez TODO1 par les lignes suivantes. Vous créez la méthode `getData` et la route « /api/values » côté serveur dans les étapes suivantes. Une URL relative est utilisée pour le point de terminaison car il doit être hébergé sur le même domaine que votre complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="5d388-p127">En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="5d388-p128">Cette méthode appelle un point de terminaison d’API Web spécifié et lui transmet le même jeton d’accès que l’application hôte Office a utilisé pour accéder à votre complément. Côté serveur, ce jeton d’accès est utilisé dans le flux « de la part de » pour obtenir un jeton d’accès à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="5d388-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="5d388-229">Vous créerez la méthode `handleServerSideErrors` à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="5d388-229">You will create the `handleServerSideErrors` method in a later step.</span></span>

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        });
    }
    ```

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="5d388-230">Création des méthodes de gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="5d388-230">Create the error-handling methods</span></span>

1. <span data-ttu-id="5d388-p129">En dessous de la méthode `getData`, ajoutez la méthode suivante. Cette méthode gérera les erreurs dans le client du complément lorsque l’hôte Office ne parviendra pas à obtenir un jeton d’accès pour le service web du complément. Ces erreurs sont signalées avec un code d’erreur, donc la méthode utilise une instruction `switch` pour les distinguer.</span><span class="sxs-lookup"><span data-stu-id="5d388-p129">Below the `getData` method, add the following method. This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service. These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```js
    function handleClientSideErrors(result) {

        switch (result.error.code) {

            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor.

            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.

            // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
            //        the user has not granted the service permission to their `profile`.

            // TODO6: Handle an unspecified error from the Office host.

            // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
            //        web service/application.

            // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
            //        before a previous call of it completed.

            // TODO9: Handle the case where the add-in does not support forcing consent.

            // TODO10: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="5d388-p130">Remplacez `TODO2` par le code suivant. L’erreur 13001 se produit si l’utilisateur n’est pas connecté, ou s’il a annulé, sans y répondre, une invite lui demandant d’indiquer un deuxième facteur d’authentification. Dans les deux cas, le code réexécute la méthode `getDataWithToken` et définit une option pour forcer une invite de connexion.</span><span class="sxs-lookup"><span data-stu-id="5d388-p130">Replace `TODO2` with the following code. Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor. In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```js
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="5d388-p131">Remplacez `TODO3` par le code suivant. L’erreur 13002 se produit lorsque la connexion ou l’octroi du consentement de l’utilisateur a été abandonné. Demandez à l’utilisateur de réessayer, mais seulement une fois.</span><span class="sxs-lookup"><span data-stu-id="5d388-p131">Replace `TODO3` with the following code. Error 13002 occurs when user's sign-in or consent was aborted. Ask the user to try again but no more than once again.</span></span>

    ```js
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. <span data-ttu-id="5d388-240">Remplacez `TODO4` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-240">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="5d388-241">L’erreur 13003 se produit si l’utilisateur est connecté avec un compte qui n’est ni un compte professionnel ni un compte scolaire, ni un compte Microsoft.</span><span class="sxs-lookup"><span data-stu-id="5d388-241">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft account.</span></span> <span data-ttu-id="5d388-242">Demandez à l’utilisateur de se déconnecter, puis de se reconnecter avec un type de compte pris en charge.</span><span class="sxs-lookup"><span data-stu-id="5d388-242">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```js
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > <span data-ttu-id="5d388-243">L’erreur 13004 n’est pas gérée dans cette méthode, car elle ne devrait se produire qu’en développement.</span><span class="sxs-lookup"><span data-stu-id="5d388-243">Error 13004 is not handled in this method because it should only occur in development.</span></span> <span data-ttu-id="5d388-244">Elle ne peut pas être résolue par du code d’exécution et il ne serait d’aucune utilité de la signaler à un utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="5d388-244">It cannot be fixed by runtime code and there would be no point in reporting it to an end user.</span></span>

1. <span data-ttu-id="5d388-245">Remplacez `TODO5` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-245">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="5d388-246">L’erreur 13005 se produit si Office n’a pas été autorisé à accéder au service web du complément ou si l’utilisateur n’a pas accordé l’autorisation de service à son `profile`.</span><span class="sxs-lookup"><span data-stu-id="5d388-246">Error 13005 occurs when Office has not been authorized to the add-in's web service or the user has not granted the service permission to their `profile`.</span></span>

    ```js
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. <span data-ttu-id="5d388-p135">Remplacez `TODO6` par le code suivant. L’erreur 13006 se produit lorsqu’une erreur non spécifiée indiquant que l’hôte est dans un état instable est survenue dans l’hôte Office. Demandez à l’utilisateur de redémarrer Office.</span><span class="sxs-lookup"><span data-stu-id="5d388-p135">Replace `TODO6` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```js
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. <span data-ttu-id="5d388-p136">Remplacez `TODO7` par le code suivant. L’erreur 13007 se produit lorsqu’un problème est survenu au niveau de l’interaction de l’hôte Office avec AAD de telle sorte que l’hôte ne peut pas obtenir de jeton d’accès pour accéder à l’application/au service Web des compléments. Il peut s’agir d’un problème temporaire de réseau. Demandez à l’utilisateur de réessayer plus tard.</span><span class="sxs-lookup"><span data-stu-id="5d388-p136">Replace `TODO7` with the following code. Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application. This may be a temporary network issue. Ask the user to try again later.</span></span>

    ```js
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. <span data-ttu-id="5d388-254">Remplacez `TODO8` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-254">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="5d388-255">L’erreur 13008 se produit lorsque l’utilisateur a déclenché une opération qui appelle `getAccessTokenAsync` avant que la fin de l’appel précédent.</span><span class="sxs-lookup"><span data-stu-id="5d388-255">Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```js
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. <span data-ttu-id="5d388-p138">Remplacez `TODO9` par le code suivant. L’erreur 13009 se produit lorsque le complément ne prend pas en charge l’obligation d’afficher une invite de consentement, mais que `getAccessTokenAsync` a été appelé avec l’option `forceConsent` définie sur `true`. Dans le cas habituel, lorsque cela se produit, le code doit automatiquement réexécuter `getAccessTokenAsync` avec l’option de consentement définie sur `false`. Toutefois, dans certains cas, l’appel de la méthode avec `forceConsent` défini sur `true` était lui-même une réponse automatique à une erreur dans un appel à la méthode avec l’option définie sur `false`. Dans ce cas, le code ne doit pas réessayer, mais il doit à la place conseiller à l’utilisateur de se déconnecter et de se reconnecter.</span><span class="sxs-lookup"><span data-stu-id="5d388-p138">Replace `TODO9` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```js
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```

1. <span data-ttu-id="5d388-261">Remplacez `TODO10` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-261">Replace `TODO10` with the following code.</span></span>

    ```js
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="5d388-p139">En dessous de la méthode `handleClientSideErrors`, ajoutez la méthode suivante. Cette méthode gérera les erreurs du service web du complément en cas de problème d’exécution du flux « de la part de » ou de problème d’obtention de données à partir de Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="5d388-p139">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```js
    function handleServerSideErrors(result) {

        // TODO11: Parse the JSON response.

        // TODO12: Handle the case where AAD asks for an additional form of authentication.

        // TODO13: Handle missing consent and scope (permission) related issues.

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="5d388-264">Remplacez `TODO11` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-264">Replace `TODO11` with the following code.</span></span> <span data-ttu-id="5d388-265">Pour la plupart des erreurs `4xx` que le service web du complément transmettra du côté client du complément, une propriété **ExceptionMessage** se trouvera dans la réponse contenant le numéro d’erreur AADSTS (Azure Active Directory Secure Token Service), ainsi que d’autres données.</span><span class="sxs-lookup"><span data-stu-id="5d388-265">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="5d388-266">Toutefois, lorsqu’AAD enverra un message au service web du complément pour demander un facteur d’authentification supplémentaire, le message contiendra une propriété **Claims** spéciale spécifiant (avec un numéro de code) le facteur supplémentaire nécessaire.</span><span class="sxs-lookup"><span data-stu-id="5d388-266">However, when AAD sends a message to the add-in's web service asking for an additional authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="5d388-267">Les API ASP.NET qui créent et envoient des réponses HTTP aux clients ne connaissent pas cette propriété **Claims**, donc ils ne l’incluent pas dans l’objet de la réponse.</span><span class="sxs-lookup"><span data-stu-id="5d388-267">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="5d388-268">Le code côté serveur que vous allez créer dans une étape ultérieure y remédiera en ajoutant manuellement la valeur **Claims** à l’objet de réponse.</span><span class="sxs-lookup"><span data-stu-id="5d388-268">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="5d388-269">Cette valeur sera dans la propriété **Message**, donc le code doit également analyser cette propriété.</span><span class="sxs-lookup"><span data-stu-id="5d388-269">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="5d388-270">Remplacez `TODO12` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-270">Replace `TODO12` with the following code.</span></span> <span data-ttu-id="5d388-271">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-271">Note about this code:</span></span>

    * <span data-ttu-id="5d388-272">L’erreur 50076 se produit lorsque Microsoft Graph exige un formulaire d’authentification supplémentaire.</span><span class="sxs-lookup"><span data-stu-id="5d388-272">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="5d388-p142">L’hôte Office dois obtenir un nouveau jeton avec la valeur **Claims** pour l’option `authChallenge`. Cela demande à AAD d’inviter l’utilisateur à accepter tous les formulaires d’authentification requis.</span><span class="sxs-lookup"><span data-stu-id="5d388-p142">The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```js
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }
    ```

1. <span data-ttu-id="5d388-275">Remplacez `TODO13` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-275">Replace `TODO13` with the following code.</span></span> <span data-ttu-id="5d388-276">Dans les prochaines étapes, vous allez remplacer les trois `TODO` dans ce code par un bloc conditionnel *interne*.</span><span class="sxs-lookup"><span data-stu-id="5d388-276">You will replace the three `TODO`s in this code with an *inner* conditional block in the next few steps.</span></span>

    ```js
    else if (exceptionMessage) {

        // TODO13A: Handle the case where consent has not been granted, or has been revoked.

        // TODO13B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO13C: Handle the case where the token that the add-in's client-side sends to it's
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. <span data-ttu-id="5d388-277">Remplacez `TODO13A` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-277">Replace `TODO13A` with the following code.</span></span> <span data-ttu-id="5d388-278">(Cette opération crée la première partie d’un bloc conditionnel *interne*.) Note sur ce code :</span><span class="sxs-lookup"><span data-stu-id="5d388-278">(This creates the first part of an *inner* conditional block.) Note about this code:</span></span>

    * <span data-ttu-id="5d388-279">L’erreur 65001 signifie que l’utilisateur a refusé de donner l’accès à Microsoft Graph (ou que l’accès a été révoqué) pour une ou plusieurs autorisations.</span><span class="sxs-lookup"><span data-stu-id="5d388-279">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span>
    * <span data-ttu-id="5d388-280">Le complément doit obtenir un nouveau jeton avec l’option `forceConsent` définie sur `true`.</span><span class="sxs-lookup"><span data-stu-id="5d388-280">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```js
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
       getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="5d388-p145">Remplacez `TODO13B` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p145">Replace `TODO13B` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5d388-p146">L’erreur 70011 a plusieurs sens. Le problème qui importe pour ce complément est lorsque cette erreur indique qu’une étendue (autorisation) non valide a été demandée ; le code vérifie alors la description complète de l’erreur, pas seulement le numéro.</span><span class="sxs-lookup"><span data-stu-id="5d388-p146">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="5d388-285">Le complément doit signaler l’erreur.</span><span class="sxs-lookup"><span data-stu-id="5d388-285">The add-in should report the error.</span></span>

    ```js
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="5d388-p147">Remplacez `TODO13C` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p147">Replace `TODO13C` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5d388-288">Le code côté serveur que vous allez créer dans une étape ultérieure enverra le message `Missing access_as_user` si l’étendue (autorisation) `access_as_user` ne se trouve pas dans le jeton d’accès que le client du complément envoie à AAD pour qu’il l’utilise dans flux « de la part de ».</span><span class="sxs-lookup"><span data-stu-id="5d388-288">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="5d388-289">Le complément doit signaler l’erreur.</span><span class="sxs-lookup"><span data-stu-id="5d388-289">The add-in should report the error.</span></span>

    ```js
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="5d388-290">Remplacez `TODO14` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-290">Replace `TODO14` with the following code.</span></span> <span data-ttu-id="5d388-291">(Cela fait partie du bloc conditionnel *externe* et doit figurer immédiatement après le crochet fermant de la structure commençant par `else if (exceptionMessage) {` et au même niveau de retrait.) Note sur ce code :</span><span class="sxs-lookup"><span data-stu-id="5d388-291">(This is part of the *outer* conditional block and should be immediately after the close bracket of the structure that begins with `else if (exceptionMessage) {` and at the same level of indentation.) Note about this code:</span></span>

    * <span data-ttu-id="5d388-292">La bibliothèque d’identité que vous allez utiliser dans le code côté serveur (Microsoft Authentication Library, MSAL) doit garantir qu’aucun jeton expiré ou non valide n’est envoyé à Microsoft Graph. Cependant, si cela se produit, l’erreur renvoyée par Microsoft Graph au service web du complément a le code `InvalidAuthenticationToken`.</span><span class="sxs-lookup"><span data-stu-id="5d388-292">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`.</span></span> <span data-ttu-id="5d388-293">Le code côté serveur que vous allez créer dans une étape ultérieure envoie ce message au client du complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-293">Server-side code you will create in a later step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="5d388-294">Dans ce cas, le complément doit recommencer l’intégralité du processus d’authentification en réinitialisant les variables de compteur et d’indicateur, puis en appelant à nouveau la méthode de gestionnaire de boutons.</span><span class="sxs-lookup"><span data-stu-id="5d388-294">In this case, the add-in should start the entire authentication process over by resetting the counter and flag variables, and then re-calling the button handler method.</span></span>

    ```js
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }
    ```

1. <span data-ttu-id="5d388-295">Remplacez `TODO15` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="5d388-295">Replace `TODO15` with the following code.</span></span>

    ```js
    else {
        logError(result);
    }
    ```

1. <span data-ttu-id="5d388-296">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-296">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="5d388-297">Code côté serveur</span><span class="sxs-lookup"><span data-stu-id="5d388-297">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="5d388-298">Configurer les intergiciels OWIN</span><span class="sxs-lookup"><span data-stu-id="5d388-298">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="5d388-299">Ouvrez le fichier Startup.cs à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="5d388-299">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="5d388-p150">Ajoutez le mot clé `partial` à la déclaration de la classe de démarrage, si ce n’est pas déjà fait. Elle doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="5d388-p150">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="5d388-p151">Ajoutez la ligne suivante dans le corps de la méthode `Configuration`. Vous créez la méthode `ConfigureAuth` dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="5d388-p151">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="5d388-304">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-304">Save and close the file.</span></span>

1. <span data-ttu-id="5d388-305">Cliquez avec le bouton droit de la souris sur le dossier **App_Start**, puis sélectionnez **Ajouter > Classe**.</span><span class="sxs-lookup"><span data-stu-id="5d388-305">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="5d388-306">Dans la boîte de dialogue **Ajouter un nouvel élément** nommez le fichier **Startup.Auth.cs**, puis cliquez sur **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="5d388-306">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="5d388-307">Raccourcissez le nom de l’espace de noms dans le nouveau fichier `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="5d388-307">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="5d388-308">Vérifiez que toutes les instructions `using` suivantes se trouvent en haut du fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-308">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="5d388-p152">Ajoutez le mot clé `partial` à la déclaration de la classe `Startup`, si ce n’est pas déjà fait. Elle doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="5d388-p152">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="5d388-p153">Ajoutez la méthode suivante à la classe `Startup`. Cette méthode spécifie comment l’intergiciel OWIN valide les jetons d’accès qui lui sont transmis à partir de la méthode `getData` dans le fichier Home.js côté client. Le processus d’autorisation est déclenché chaque fois qu’un point de terminaison Web API décoré avec l’attribut `[Authorize]` est appelé.</span><span class="sxs-lookup"><span data-stu-id="5d388-p153">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="5d388-p154">Remplacez TODO3 par les lignes suivantes. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p154">Replace the TODO3 with the following. Note about this code:</span></span>

    * <span data-ttu-id="5d388-316">Le code demande à OWIN de s’assurer que l’audience et l’émetteur du jeton spécifiés dans le jeton d’accès qui provient de l’hôte Office (et est transmis par l’appel côté client de `getData`) doivent correspondre aux valeurs spécifiées dans le fichier web.config.</span><span class="sxs-lookup"><span data-stu-id="5d388-316">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="5d388-p155">Le réglage de `SaveSigninToken` sur `true` fait qu’OWIN enregistre le jeton brut à partir de l’hôte Office. Le complément en a besoin pour obtenir un jeton d’accès à Microsoft Graph avec le flux « de la part de ».</span><span class="sxs-lookup"><span data-stu-id="5d388-p155">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.</span></span>
    * <span data-ttu-id="5d388-p156">Les étendues ne sont pas validées par l’intergiciel OWIN. Les étendues du jeton d’accès, qui doivent inclure `access_as_user`, sont validées dans le contrôleur.</span><span class="sxs-lookup"><span data-stu-id="5d388-p156">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="5d388-p157">Remplacez TODO4 par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p157">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="5d388-323">La méthode `UseOAuthBearerAuthentication` est appelée au lieu de la méthode `UseWindowsAzureActiveDirectoryBearerAuthentication` plus courante, car cette dernière n’est pas compatible avec le point de terminaison Azure AD V2.</span><span class="sxs-lookup"><span data-stu-id="5d388-323">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="5d388-324">L’URL de découverte transmise à la méthode correspond à l’endroit où l’intergiciel OWIN obtient les instructions permettant d’obtenir la clé requise pour vérifier la signature sur le jeton d’accès reçu de l’hôte Office.</span><span class="sxs-lookup"><span data-stu-id="5d388-324">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="5d388-325">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-325">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="5d388-326">Créer le contrôleur /api/values</span><span class="sxs-lookup"><span data-stu-id="5d388-326">Create the /api/values controller</span></span>

1. <span data-ttu-id="5d388-327">Ouvrez le fichier **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="5d388-327">Open the file **Controllers\ValueController.cs**.</span></span>

1. <span data-ttu-id="5d388-328">Vérifiez que les instructions `using` suivantes se trouvent en haut du fichier.</span><span class="sxs-lookup"><span data-stu-id="5d388-328">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

1. <span data-ttu-id="5d388-p158">Juste au-dessus de la ligne qui déclare `ValuesController`, ajoutez l’attribut `[Authorize]`. Cela permet de s’assurer que votre complément exécutera le processus d’autorisation que vous avez configuré dans la dernière procédure chaque fois qu’une méthode de contrôleur est appelée. Seuls les appelants avec un jeton d’accès valide à votre complément peuvent ainsi appeler les méthodes du contrôleur.</span><span class="sxs-lookup"><span data-stu-id="5d388-p158">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5d388-p159">Un service d’API Web MVC ASP.NET en production doit avoir une logique personnalisée pour le flux « de la part de » dans une ou plusieurs classes **FilterAttribute** personnalisées. Cet exemple pédagogique place la logique dans le contrôleur principal afin que l’intégralité du flux de la logique d’extraction de données et d’autorisation puisse être facilement suivie. De plus, l’exemple est cohérent avec les exemples de modèle d’autorisation dans [Exemples Azure](https://github.com/Azure-Samples/).</span><span class="sxs-lookup"><span data-stu-id="5d388-p159">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom **FilterAttribute** classes. This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed. This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>

1. <span data-ttu-id="5d388-p160">Ajoutez la méthode suivante à `ValuesController`. Vous remarquerez que la valeur renvoyée est `Task<HttpResponseMessage>` et non `Task<IEnumerable<string>>`, laquelle serait plus courante pour une méthode `GET api/values`. Il s’agit d’un effet secondaire du fait que notre logique d’autorisation personnalisée se trouvera dans le contrôleur : certaines conditions d’erreur de cette logique nécessitent qu’un objet Réponse HTTP soit envoyé au client du complément.</span><span class="sxs-lookup"><span data-stu-id="5d388-p160">Add the following method to the `ValuesController`. Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method. This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

1. <span data-ttu-id="5d388-338">Remplacez `TODO1` par le code suivant pour confirmer que les étendues spécifiées dans le jeton incluent `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="5d388-338">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > <span data-ttu-id="5d388-339">Vous ne pouvez utiliser l’étendue `access_as_user` que pour autoriser l’API qui gère le flux « de la part de » pour les compléments Office. D’autres API dans votre service peuvent avoir leurs propres exigences d’étendue.</span><span class="sxs-lookup"><span data-stu-id="5d388-339">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="5d388-340">Cela permet de limiter ce à quoi donnent accès les jetons acquis par Office.</span><span class="sxs-lookup"><span data-stu-id="5d388-340">This limits what can be accessed with the tokens that Office acquires.</span></span>

1. <span data-ttu-id="5d388-p162">Remplacez `TODO2` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p162">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="5d388-343">Ce code transforme le jeton d’accès brut reçu de l’hôte Office en objet `UserAssertion` qui sera transmis à une autre méthode.</span><span class="sxs-lookup"><span data-stu-id="5d388-343">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="5d388-p163">Votre complément ne joue plus le rôle d’une ressource (ou audience) à laquelle l’hôte Office et l’utilisateur doivent accéder. Désormais, il est lui-même un client qui a besoin d’accéder à Microsoft Graph. `ConfidentialClientApplication` est l’objet de « contexte client » MSAL.</span><span class="sxs-lookup"><span data-stu-id="5d388-p163">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="5d388-p164">Le troisième paramètre du constructeur `ConfidentialClientApplication` est une URL de redirection qui n’est pas utilisée dans le flux « de la part de », mais il est recommandé d’utiliser l’URL correcte. Les quatrième et cinquième paramètres peuvent être utilisés pour définir un magasin permanent qui permettrait la réutilisation des jetons non expirés entre différentes sessions avec le complément. Cet exemple n’implémente pas un stockage permanent.</span><span class="sxs-lookup"><span data-stu-id="5d388-p164">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="5d388-p165">MSAL requiert les étendues `openid` et `offline_access` pour fonctionner, mais il génère une erreur si votre code les demande de façon redondante. Il génère également une erreur si votre code demande `profile`, qui est utilisé uniquement lorsque l’application Office hôte obtient le jeton pour l’application web de votre complément. Seul `Files.Read.All` est demandé explicitement.</span><span class="sxs-lookup"><span data-stu-id="5d388-p165">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

1. <span data-ttu-id="5d388-p166">Remplacez `TODO3` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p166">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5d388-p167">La méthode `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` recherchera tout d’abord dans le cache MSAL, c’est-à-dire en mémoire, un jeton d’accès correspondant. Uniquement s’il n’existe pas, elle lance le flux « de la part de » avec le point de terminaison Azure AD V2.</span><span class="sxs-lookup"><span data-stu-id="5d388-p167">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="5d388-357">Si une authentification multifacteur est requise par la ressource MS Graph et si l’utilisateur ne l'a pas encore fournie, AAD lève une exception qui contient une propriété de revendication.</span><span class="sxs-lookup"><span data-stu-id="5d388-357">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="5d388-p168">La valeur de la propriété Claims doit être transmise au client qui la transmettra à son tour à l’hôte Office, qui l’inclura alors dans une demande de nouveau jeton. AAD demandera à l’utilisateur d’accepter tous les formulaires d’authentification requis.</span><span class="sxs-lookup"><span data-stu-id="5d388-p168">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="5d388-360">Les exceptions qui ne sont pas de type `MsalServiceException` ne sont intentionnellement pas capturées afin d’être propagées au client sous la forme de messages `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="5d388-360">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

1. <span data-ttu-id="5d388-p169">Remplacez `TODO3a` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p169">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5d388-p170">Si l’authentification multifacteur est requise par la ressource MS Graph et que l’utilisateur ne l'a pas encore fournie, AAD renvoie « 400 - Demande incorrecte » avec l’erreur AADSTS50076 et une propriété **Claims**. MSAL génère une exception **MsalUiRequiredException** (qui hérite de **MsalServiceException**) avec ces informations.</span><span class="sxs-lookup"><span data-stu-id="5d388-p170">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="5d388-p171">La valeur de la propriété **Claims** doit être transmise au client qui doit la transmettre à son tour à l’hôte Office, qui l’inclut alors dans une demande de nouveau jeton. AAD demandera à l’utilisateur d’accepter tous les formulaires d’authentification requis.</span><span class="sxs-lookup"><span data-stu-id="5d388-p171">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="5d388-p172">Les API qui créent des réponses HTTP à partir d’exceptions ne connaissent pas la propriété **Claims**, donc ils ne l’incluent pas dans l’objet de la réponse. Nous devons créer manuellement un message qui l’inclut. Une propriété **Message** personnalisé, cependant, bloque la création d’une propriété **ExceptionMessage**, afin que la seule façon de communiquer l’ID d’erreur `AADSTS50076` au client est de l’ajouter à la propriété **Message** personnalisée. JavaScript dans le client devra découvrir si une réponse a une propriété **Message** ou **ExceptionMessage**, afin qu’il sache laquelle lire.</span><span class="sxs-lookup"><span data-stu-id="5d388-p172">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="5d388-371">Le message personnalisé est au format JSON pour que le code JavaScript côté client puisse l’analyser avec des méthodes d’objet `JSON` connues.</span><span class="sxs-lookup"><span data-stu-id="5d388-371">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="5d388-p173">Vous créerez la méthode `SendErrorToClient` à une étape ultérieure. Son deuxième paramètre est un objet **Exception**. Dans ce cas, le code transmet `null` car même l’objet **Exception** bloque l’inclusion de la propriété **Message** dans la réponse HTTP qui est générée.</span><span class="sxs-lookup"><span data-stu-id="5d388-p173">You will create the `SendErrorToClient` method in a later step. It's second parameter is an **Exception** object. In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="5d388-p174">Remplacez `TODO3b` et `TODO3c` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p174">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="5d388-p175">Si l’appel à AAD contenait au moins une étendue (autorisation) pour laquelle ni l’utilisateur, ni un administrateur client a consenti (ou pour laquelle le consentement a été révoqué) : AAD renverra « 400 Demande incorrecte » avec l’erreur `AADSTS65001`. MSAL génère une exception **MsalUiRequiredException** avec ces informations. Le client doit de nouveau appeler `getAccessTokenAsync` avec l’option `{ forceConsent: true }`.</span><span class="sxs-lookup"><span data-stu-id="5d388-p175">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked). AAD will return "400 Bad Request" with error `AADSTS65001`. MSAL throws a **MsalUiRequiredException** with this information. The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="5d388-p176">Si l’appel à AAD contenait au moins une étendue non reconnue par AAD, AAD renvoie « 400 Demande incorrecte » avec l’erreur `AADSTS70011`. MSAL génère une exception **MsalUiRequiredException** avec ces informations. Le client doit informer l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5d388-p176">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`. MSAL throws a **MsalUiRequiredException** with this information. The client should inform the user.</span></span>
    *  <span data-ttu-id="5d388-384">La description entière est incluse, car l’erreur 70011 est renvoyée dans d’autres conditions et elle doit être gérée dans ce complément uniquement lorsqu’elle indique une étendue non valide.</span><span class="sxs-lookup"><span data-stu-id="5d388-384">The entire description is included because 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="5d388-p177">L’objet **MsalUiRequiredException** est transmis à `SendErrorToClient`. Cela permet de garantir qu’une propriété **ExceptionMessage** qui contient les informations d’erreur est incluse dans la réponse HTTP.</span><span class="sxs-lookup"><span data-stu-id="5d388-p177">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="5d388-387">Il n’y a aucun message personnalisé, donc `null` est transmis en tant que troisième paramètre.</span><span class="sxs-lookup"><span data-stu-id="5d388-387">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="5d388-p178">Remplacez `TODO3d` par le code suivant. Vous remarquerez que le code génère de nouveau l’exception au lieu de la relayer dans une réponse HTTP personnalisée avec **HttpStatusCode.Forbidden** (401). L’effet de cette opération est l’envoi par ASP.NET de sa propre réponse HTTP avec le statut « Erreur serveur 500 ».</span><span class="sxs-lookup"><span data-stu-id="5d388-p178">Replace `TODO3d` with the following code. Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401). The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

1. <span data-ttu-id="5d388-p179">Remplacez `TODO4` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p179">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="5d388-p180">Les classes `GraphApiHelper` et `ODataHelper` sont définies dans les fichiers du dossier **Helpers**. La classe `OneDriveItem` est définie dans un fichier du dossier **Models**. La description détaillée de ces classes n’est pas pertinente pour l’autorisation ou l’authentification unique, elle est donc hors de portée de cet article.</span><span class="sxs-lookup"><span data-stu-id="5d388-p180">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="5d388-396">Vous pouvez améliorer les performances en ne demandant à Microsoft Graph que les données réellement requises. Ainsi, le code utilise le paramètre de requête `$select` pour spécifier que nous ne souhaitons que la propriété de nom, et le paramètre `$top` pour spécifier que nous ne voulons que les trois premiers noms de fichier ou de dossier.</span><span class="sxs-lookup"><span data-stu-id="5d388-396">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a `$select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="5d388-p181">Si le jeton envoyé à Microsoft Graph n’est pas valide, Microsoft Graph envoie l’erreur « 401 accès non autorisé » avec le code « InvalidAuthenticationToken ». ASP.NET génère ensuite une exception **RuntimeBinderException**. C’est également ce qu’il se passe lorsque le jeton a expiré, bien que MSAL doive l’empêcher.</span><span class="sxs-lookup"><span data-stu-id="5d388-p181">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken". ASP.NET then throws a **RuntimeBinderException**. This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);
    }
    ```

1. <span data-ttu-id="5d388-p182">Remplacez `TODO5` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5d388-p182">Replace `TODO5` with the following. Note about this code:</span></span>

    * <span data-ttu-id="5d388-p183">Bien que le code ci-dessus demande uniquement la propriété *name* des éléments OneDrive, Microsoft Graph comporte toujours la propriété *eTag* pour les éléments OneDrive. Pour réduire la charge utile envoyée au client, le code ci-dessous reconstruit les résultats avec uniquement les noms d’élément.</span><span class="sxs-lookup"><span data-stu-id="5d388-p183">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="5d388-404">La liste des trois fichiers et dossiers OneDrive est envoyée au client en tant que réponse HTTP « 200 OK ».</span><span class="sxs-lookup"><span data-stu-id="5d388-404">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames);
    return response;
    ```

1. <span data-ttu-id="5d388-p184">Au-dessous de la méthode Get, ajoutez la méthode suivante. Tenez compte des informations suivantes sur ce code :</span><span class="sxs-lookup"><span data-stu-id="5d388-p184">Below the Get method, add the following method. About this code note:</span></span>  

    * <span data-ttu-id="5d388-407">La méthode communique au client les informations sur une exception côté serveur.</span><span class="sxs-lookup"><span data-stu-id="5d388-407">The method relays to the client information about a server-side exception.</span></span>
    * <span data-ttu-id="5d388-408">Si l’exception d’origine est transmise à la méthode, le constructeur HttpError inclura les informations de l’objet d’exception dans une propriété **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="5d388-408">If the original exception is passed to the method, then the HttpError constructor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="5d388-409">Si `null` est transmis pour l’exception, le constructeur HttpError inclura le paramètre de message dans une propriété **Message** et aucune propriété **ExceptionMessage** ne sera présente.</span><span class="sxs-lookup"><span data-stu-id="5d388-409">If `null` is passed for the exception, then the HttpError constructor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }
    ```

## <a name="run-the-add-in"></a><span data-ttu-id="5d388-410">Exécution du complément</span><span class="sxs-lookup"><span data-stu-id="5d388-410">Run the add-in</span></span>

1. <span data-ttu-id="5d388-411">Assurez-vous que vous disposez de fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.</span><span class="sxs-lookup"><span data-stu-id="5d388-411">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="5d388-p185">Dans Visual Studio, appuyez sur F5. PowerPoint s’ouvre et un groupe **SSO ASP.NET** se trouve sur le ruban **Accueil**.</span><span class="sxs-lookup"><span data-stu-id="5d388-p185">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="5d388-414">Appuyez sur le bouton **Afficher le complément** dans ce groupe pour voir l’interface utilisateur du complément dans le volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="5d388-414">Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.</span></span>

1. <span data-ttu-id="5d388-p186">Appuyez sur le bouton **Obtenir mes fichiers à partir de OneDrive**. Si vous n’êtes pas connecté à Office, vous serez invité à vous connecter.</span><span class="sxs-lookup"><span data-stu-id="5d388-p186">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5d388-p187">Si vous étiez précédemment connecté à Office avec un ID différent, et si certaines applications Office sont toujours ouvertes, Office ne changera pas systématiquement votre identifiant même s’il semble l’avoir fait dans PowerPoint. Dans ce cas, l’appel vers Microsoft Graph peut échouer, ou des données de l’ID précédent peuvent être renvoyées. Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir mes fichiers à partir de OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="5d388-p187">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="5d388-p188">Une fois que vous êtes connecté, la liste de vos fichiers et dossiers dans OneDrive s’affiche sous le bouton. Cette opération peut prendre plus de 15 secondes, surtout la première fois.</span><span class="sxs-lookup"><span data-stu-id="5d388-p188">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
