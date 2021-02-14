---
title: Personnaliser votre complément compatible avec l’authentification unique Node.js
description: En savoir plus sur la personnalisation du module de personnalisation de LSO que vous avez créé avec le générateur Yeoman.
ms.date: 02/01/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 174df5e58e794b94b02025bd90a65f5ae8e26d44
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234169"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a><span data-ttu-id="412ba-103">Personnaliser votre complément compatible avec l’authentification unique Node.js</span><span class="sxs-lookup"><span data-stu-id="412ba-103">Customize your Node.js SSO-enabled add-in</span></span>

> [!IMPORTANT]
> <span data-ttu-id="412ba-104">Cet article s’appuie sur le compl?ment sso-enabled qui est créé en compl?tant le démarrage rapide de l' [sign-on unique (SSO).](sso-quickstart.md)</span><span class="sxs-lookup"><span data-stu-id="412ba-104">This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md).</span></span> <span data-ttu-id="412ba-105">Veuillez effectuer le démarrage rapide avant de lire cet article.</span><span class="sxs-lookup"><span data-stu-id="412ba-105">Please complete the quick start before reading this article.</span></span>

<span data-ttu-id="412ba-106">Le [](sso-quickstart.md) démarrage rapide de l' cesso crée un add-in ssO qui obtient les informations de profil de l’utilisateur et les écrit dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="412ba-106">The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document or message.</span></span> <span data-ttu-id="412ba-107">Dans cet article, vous allez passer en revue le processus de mise à jour du add-in que vous avez créé avec le générateur Yeoman dans le démarrage rapide de l’eoso, pour ajouter de nouvelles fonctionnalités qui nécessitent différentes autorisations.</span><span class="sxs-lookup"><span data-stu-id="412ba-107">In this article, you'll walk through the process of updating the add-in that you created with the Yeoman generator in the SSO quick start, to add new functionality that requires different permissions.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="412ba-108">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="412ba-108">Prerequisites</span></span>

- <span data-ttu-id="412ba-109">Un add-in Office que vous avez créé en suivant les instructions du démarrage rapide de [l' cesso.](sso-quickstart.md)</span><span class="sxs-lookup"><span data-stu-id="412ba-109">An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).</span></span>

- <span data-ttu-id="412ba-110">Au moins quelques fichiers et dossiers stockés sur OneDrive Entreprise dans votre abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="412ba-110">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

- <span data-ttu-id="412ba-111">[Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="412ba-111">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a><span data-ttu-id="412ba-112">Passer en revue le contenu du projet</span><span class="sxs-lookup"><span data-stu-id="412ba-112">Review contents of the project</span></span>

<span data-ttu-id="412ba-113">Commençons par un examen rapide du projet de add-in que vous avez précédemment créé avec le [générateur Yeoman.](sso-quickstart.md)</span><span class="sxs-lookup"><span data-stu-id="412ba-113">Let's begin with a quick review of the add-in project that you previously [created with the Yeoman generator](sso-quickstart.md).</span></span>

> [!NOTE]
> <span data-ttu-id="412ba-114">À des endroits où cet article fait référence à des fichiers de script à l’aide de l’extension de fichier **.js,** supposez plutôt l’extension de fichier **.ts** si votre projet a été créé avec TypeScript.</span><span class="sxs-lookup"><span data-stu-id="412ba-114">In places where this article references script files using **.js** file extension, assume the **.ts** file extension instead if your project was created with TypeScript.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a><span data-ttu-id="412ba-115">Ajouter de nouvelles fonctionnalités</span><span class="sxs-lookup"><span data-stu-id="412ba-115">Add new functionality</span></span>

<span data-ttu-id="412ba-116">Le add-in que vous avez créé avec le démarrage rapide de l' cesso utilise Microsoft Graph pour obtenir les informations de profil de l’utilisateur et écrit ces informations dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="412ba-116">The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document or message.</span></span> <span data-ttu-id="412ba-117">Nous allons modifier la fonctionnalité du add-in de telle façon qu’il obtient les noms des 10 principaux fichiers et dossiers du OneDrive Entreprise de l’utilisateur et écrit ces informations dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="412ba-117">Let's change the add-in's functionality such that it gets the names of the top 10 files and folders from the signed-in user's OneDrive for Business and writes that information to the document or message.</span></span> <span data-ttu-id="412ba-118">L’activation de cette nouvelle fonctionnalité nécessite la mise à jour des autorisations d’application dans Azure et la mise à jour du code dans le projet de add-in.</span><span class="sxs-lookup"><span data-stu-id="412ba-118">Enabling this new functionality requires updating app permissions in Azure and updating code within the add-in project.</span></span>

### <a name="update-app-permissions-in-azure"></a><span data-ttu-id="412ba-119">Mettre à jour les autorisations d’application dans Azure</span><span class="sxs-lookup"><span data-stu-id="412ba-119">Update app permissions in Azure</span></span>

<span data-ttu-id="412ba-120">Pour que le module puisse lire correctement le contenu de OneDrive Entreprise de l’utilisateur, ses informations d’inscription d’application dans Azure doivent être mises à jour avec les autorisations appropriées.</span><span class="sxs-lookup"><span data-stu-id="412ba-120">Before the add-in can successfully read the contents of the user's OneDrive for Business, its app registration information in Azure must be updated with the appropriate permissions.</span></span> <span data-ttu-id="412ba-121">Pour accorder à l’application **l’autorisation Files.Read.All** et révoquer l’autorisation **User.Read,** qui n’est plus nécessaire, complétez les étapes suivantes.</span><span class="sxs-lookup"><span data-stu-id="412ba-121">Complete the following steps to grant the app the **Files.Read.All** permission and revoke the **User.Read** permission, which is no longer needed.</span></span>

1. <span data-ttu-id="412ba-122">Accédez au [portail Azure et](https://ms.portal.azure.com/#home) **connectez-vous à l’aide** de vos informations d’identification d’administrateur Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="412ba-122">Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and **sign in using your Microsoft 365 administrator credentials**.</span></span>

2. <span data-ttu-id="412ba-123">Accédez à la page **Inscriptions des applications.**</span><span class="sxs-lookup"><span data-stu-id="412ba-123">Navigate to the **App registrations** page.</span></span>
    > [!TIP]
    > <span data-ttu-id="412ba-124">Pour ce faire, vous  pouvez choisir la vignette Inscriptions de l’application sur la page d’accueil Azure ou à l’aide de la zone de recherche de la page d’accueil pour rechercher et choisir les inscriptions **d’applications.**</span><span class="sxs-lookup"><span data-stu-id="412ba-124">You can do this either by choosing the **App registrations** tile on the Azure home page or by using the search box on the home page to find and choose **App registrations**.</span></span>

3. <span data-ttu-id="412ba-125">Dans la page **Inscriptions de l’application,** choisissez l’application que vous avez créée lors du démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="412ba-125">On the **App registrations** page, choose the app that you created during the quick start.</span></span>
    > [!TIP]
    > <span data-ttu-id="412ba-126">Le **nom d’affichage** de l’application correspond au nom de la application que vous avez spécifié lors de la création du projet avec le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="412ba-126">The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.</span></span>

4. <span data-ttu-id="412ba-127">Dans la page vue d’ensemble  de l’application, choisissez les **autorisations d’API** sous le titre Gérer sur le côté gauche de la page.</span><span class="sxs-lookup"><span data-stu-id="412ba-127">From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.</span></span>

5. <span data-ttu-id="412ba-128">Dans la **ligne User.Read** de la table d’autorisations, choisissez les sélections, puis sélectionnez Révoquer le consentement de l’administrateur dans le menu qui s’affiche. </span><span class="sxs-lookup"><span data-stu-id="412ba-128">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Revoke admin consent** from the menu that appears.</span></span>

6. <span data-ttu-id="412ba-129">Sélectionnez **le bouton Oui,** supprimer en réponse à l’invite qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="412ba-129">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

7. <span data-ttu-id="412ba-130">Dans la **ligne User.Read** du tableau des autorisations,  choisissez les sélections, puis sélectionnez Supprimer l’autorisation du menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="412ba-130">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Remove permission** from the menu that appears.</span></span>

8. <span data-ttu-id="412ba-131">Sélectionnez **le bouton Oui,** supprimer en réponse à l’invite qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="412ba-131">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

9. <span data-ttu-id="412ba-132">Cliquez sur le bouton **Ajouter une autorisation**.</span><span class="sxs-lookup"><span data-stu-id="412ba-132">Select the **Add a permission** button.</span></span>

10. <span data-ttu-id="412ba-133">Dans le panneau qui s’ouvre, **choisissez Microsoft Graph,** puis les **autorisations déléguées.**</span><span class="sxs-lookup"><span data-stu-id="412ba-133">On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

11. <span data-ttu-id="412ba-134">Dans le panneau **Demander des autorisations d’API** :</span><span class="sxs-lookup"><span data-stu-id="412ba-134">On the **Request API permissions** panel:</span></span>

    <span data-ttu-id="412ba-135">a.</span><span class="sxs-lookup"><span data-stu-id="412ba-135">a.</span></span> <span data-ttu-id="412ba-136">Sous **Fichiers,** **sélectionnez Files.Read.All**.</span><span class="sxs-lookup"><span data-stu-id="412ba-136">Under **Files**, select **Files.Read.All**.</span></span>

    <span data-ttu-id="412ba-137">b.</span><span class="sxs-lookup"><span data-stu-id="412ba-137">b.</span></span> <span data-ttu-id="412ba-138">Sélectionnez **le bouton Ajouter des autorisations** en bas du panneau pour enregistrer ces modifications d’autorisations.</span><span class="sxs-lookup"><span data-stu-id="412ba-138">Select the **Add permissions** button at the bottom of the panel to save these permissions changes.</span></span>

12. <span data-ttu-id="412ba-139">Sélectionnez le **bouton Accorder le consentement de l’administrateur pour [nom du client].**</span><span class="sxs-lookup"><span data-stu-id="412ba-139">Select the **Grant admin consent for [tenant name]** button.</span></span>

13. <span data-ttu-id="412ba-140">Sélectionnez **le bouton** Oui en réponse à l’invite qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="412ba-140">Select the **Yes** button in response to the prompt that's displayed.</span></span>

### <a name="update-code-in-the-add-in-project"></a><span data-ttu-id="412ba-141">Mettre à jour le code dans le projet de add-in</span><span class="sxs-lookup"><span data-stu-id="412ba-141">Update code in the add-in project</span></span>

<span data-ttu-id="412ba-142">Pour permettre au add-in de lire le contenu du OneDrive Entreprise de l’utilisateur, vous devez :</span><span class="sxs-lookup"><span data-stu-id="412ba-142">To enable the add-in to read contents of the signed-in user's OneDrive for Business, you'll need to:</span></span>

- <span data-ttu-id="412ba-143">Mettez à jour le code qui fait référence à l’URL, aux paramètres et à l’étendue d’accès requis de Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="412ba-143">Update the code that references the Microsoft Graph URL, parameters, and required access scope.</span></span>

- <span data-ttu-id="412ba-144">Mettez à jour le code qui définit l’interface utilisateur du volet Des tâches, afin qu’il décrive avec précision les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="412ba-144">Update the code that defines the task pane UI, so that it accurately describes the new functionality.</span></span>

- <span data-ttu-id="412ba-145">Mettez à jour le code qui analyse la réponse de Microsoft Graph et l’écrit dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="412ba-145">Update the code that parses the response from Microsoft Graph and writes it to the document or message.</span></span>

<span data-ttu-id="412ba-146">Les étapes suivantes décrivent ces mises à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-146">The following steps describe these updates.</span></span>

### <a name="changes-required-for-any-type-of-add-in"></a><span data-ttu-id="412ba-147">Modifications requises pour n’importe quel type de add-in</span><span class="sxs-lookup"><span data-stu-id="412ba-147">Changes required for any type of add-in</span></span>

<span data-ttu-id="412ba-148">Pour modifier l’URL, les paramètres et l’étendue d’accès de Microsoft Graph et mettre à jour l’interface utilisateur du volet Des tâches, complétez les étapes suivantes pour votre application.</span><span class="sxs-lookup"><span data-stu-id="412ba-148">Complete the following steps for your add-in, to change the Microsoft Graph URL, parameters, and access scope, and update the task pane UI.</span></span> <span data-ttu-id="412ba-149">Ces étapes sont les mêmes, quelle que soit l’application Office ciblée par votre application.</span><span class="sxs-lookup"><span data-stu-id="412ba-149">These steps are the same, regardless of which Office application your add-in targets.</span></span>

1. <span data-ttu-id="412ba-150">Dans **le ./. Fichier ENV** :</span><span class="sxs-lookup"><span data-stu-id="412ba-150">In the **./.ENV** file:</span></span>

    <span data-ttu-id="412ba-151">a.</span><span class="sxs-lookup"><span data-stu-id="412ba-151">a.</span></span> <span data-ttu-id="412ba-152">Remplacez `GRAPH_URL_SEGMENT=/me` par ce qui suit : `GRAPH_URL_SEGMENT=/me/drive/root/children`</span><span class="sxs-lookup"><span data-stu-id="412ba-152">Replace `GRAPH_URL_SEGMENT=/me` with the following: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span></span>

    <span data-ttu-id="412ba-153">b.</span><span class="sxs-lookup"><span data-stu-id="412ba-153">b.</span></span> <span data-ttu-id="412ba-154">Remplacez `QUERY_PARAM_SEGMENT=` par ce qui suit : `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span><span class="sxs-lookup"><span data-stu-id="412ba-154">Replace `QUERY_PARAM_SEGMENT=` with the following: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span></span>

    <span data-ttu-id="412ba-155">c.</span><span class="sxs-lookup"><span data-stu-id="412ba-155">c.</span></span> <span data-ttu-id="412ba-156">Remplacez `SCOPE=User.Read` par ce qui suit : `SCOPE=Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="412ba-156">Replace `SCOPE=User.Read` with the following: `SCOPE=Files.Read.All`</span></span>

2. <span data-ttu-id="412ba-157">Dans **./manifest.xml**, recherchez la ligne vers la fin du fichier et remplacez-la `<Scope>User.Read</Scope>` par la `<Scope>Files.Read.All</Scope>` ligne.</span><span class="sxs-lookup"><span data-stu-id="412ba-157">In **./manifest.xml**, find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>`.</span></span>

3. <span data-ttu-id="412ba-158">Dans **./src/helpers/fallbackauthdialog.js** (ou **dans ./src/helpers/fallbackauthdialog.ts** pour un projet TypeScript), recherchez la chaîne et remplacez-la par la chaîne définie comme suit `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All` `requestObj` :</span><span class="sxs-lookup"><span data-stu-id="412ba-158">In **./src/helpers/fallbackauthdialog.js** (or in **./src/helpers/fallbackauthdialog.ts** for a TypeScript project), find the string `https://graph.microsoft.com/User.Read` and replace it with the string `https://graph.microsoft.com/Files.Read.All`, such that `requestObj` is defined as follows:</span></span>

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. <span data-ttu-id="412ba-159">Dans **./src/taskpane/taskpane.html**, recherchez l’élément et mettez à jour le texte dans cet élément pour décrire les nouvelles fonctionnalités `<section class="ms-firstrun-instructionstep__header">` du module.</span><span class="sxs-lookup"><span data-stu-id="412ba-159">In **./src/taskpane/taskpane.html**, find the element `<section class="ms-firstrun-instructionstep__header">` and update the text within that element to describe the add-in's new functionality.</span></span>

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. <span data-ttu-id="412ba-160">Dans **./src/taskpane/taskpane.html**, recherchez et remplacez les deux occurrences de la chaîne `Get My User Profile Information` par la chaîne `Read my OneDrive for Business` .</span><span class="sxs-lookup"><span data-stu-id="412ba-160">In **./src/taskpane/taskpane.html**, find and replace both occurrences of the string `Get My User Profile Information` with the string `Read my OneDrive for Business`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. <span data-ttu-id="412ba-161">Dans **./src/taskpane/taskpane.html**, recherchez et remplacez la chaîne `Your user profile information will be displayed in the document.` par la chaîne `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .</span><span class="sxs-lookup"><span data-stu-id="412ba-161">In **./src/taskpane/taskpane.html**, find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. <span data-ttu-id="412ba-162">Mettez à jour le code qui analyse la réponse de Microsoft Graph et l’écrit dans le document ou le message, en suivant les instructions de la section qui correspond à votre type de add-in :</span><span class="sxs-lookup"><span data-stu-id="412ba-162">Update the code that parses the response from Microsoft Graph and writes it to the document or message, by following guidance in the section that corresponds to your type of add-in:</span></span>

    - [<span data-ttu-id="412ba-163">Modifications requises pour un add-in Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-163">Changes required for an Excel add-in (JavaScript)</span></span>](#changes-required-for-an-excel-add-in-javascript)
    - [<span data-ttu-id="412ba-164">Modifications requises pour un add-in Excel (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-164">Changes required for an Excel add-in (TypeScript)</span></span>](#changes-required-for-an-excel-add-in-typescript)
    - [<span data-ttu-id="412ba-165">Modifications requises pour un add-in Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-165">Changes required for an Outlook add-in (JavaScript)</span></span>](#changes-required-for-an-outlook-add-in-javascript)
    - [<span data-ttu-id="412ba-166">Modifications requises pour un add-in Outlook (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-166">Changes required for an Outlook add-in (TypeScript)</span></span>](#changes-required-for-an-outlook-add-in-typescript)
    - [<span data-ttu-id="412ba-167">Modifications requises pour un add-in PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-167">Changes required for a PowerPoint add-in (JavaScript)</span></span>](#changes-required-for-a-powerpoint-add-in-javascript)
    - [<span data-ttu-id="412ba-168">Modifications requises pour un add-in PowerPoint (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-168">Changes required for a PowerPoint add-in (TypeScript)</span></span>](#changes-required-for-a-powerpoint-add-in-typescript)
    - [<span data-ttu-id="412ba-169">Modifications requises pour un add-in Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-169">Changes required for a Word add-in (JavaScript)</span></span>](#changes-required-for-a-word-add-in-javascript)
    - [<span data-ttu-id="412ba-170">Modifications requises pour un add-in Word (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-170">Changes required for a Word add-in (TypeScript)</span></span>](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a><span data-ttu-id="412ba-171">Modifications requises pour un add-in Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-171">Changes required for an Excel add-in (JavaScript)</span></span>

<span data-ttu-id="412ba-172">Si votre add-in est un add-in Excel créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="412ba-172">If your add-in is an Excel add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="412ba-173">Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-173">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="412ba-174">Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-174">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="412ba-175">Recherchez `writeDataToExcel` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-175">Find the `writeDataToExcel` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="412ba-176">Supprimez la `writeDataToOutlook` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-176">Delete the `writeDataToOutlook` function.</span></span>

5. <span data-ttu-id="412ba-177">Supprimez la `writeDataToPowerPoint` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-177">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="412ba-178">Supprimez la `writeDataToWord` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-178">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="412ba-179">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-179">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-excel-add-in-typescript"></a><span data-ttu-id="412ba-180">Modifications requises pour un add-in Excel (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-180">Changes required for an Excel add-in (TypeScript)</span></span>

<span data-ttu-id="412ba-181">Si votre add-in est un module excel créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts,** recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :</span><span class="sxs-lookup"><span data-stu-id="412ba-181">If your add-in is an Excel add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }

    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

<span data-ttu-id="412ba-182">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-182">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-javascript"></a><span data-ttu-id="412ba-183">Modifications requises pour un add-in Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-183">Changes required for an Outlook add-in (JavaScript)</span></span>

<span data-ttu-id="412ba-184">Si votre add-in est un add-in Outlook créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="412ba-184">If your add-in is an Outlook add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="412ba-185">Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-185">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="412ba-186">Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-186">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="412ba-187">Recherchez `writeDataToOutlook` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-187">Find the `writeDataToOutlook` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. <span data-ttu-id="412ba-188">Supprimez la `writeDataToExcel` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-188">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="412ba-189">Supprimez la `writeDataToPowerPoint` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-189">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="412ba-190">Supprimez la `writeDataToWord` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-190">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="412ba-191">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-191">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-typescript"></a><span data-ttu-id="412ba-192">Modifications requises pour un add-in Outlook (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-192">Changes required for an Outlook add-in (TypeScript)</span></span>

<span data-ttu-id="412ba-193">Si votre add-in est un add-in Outlook qui a été créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts**, recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :</span><span class="sxs-lookup"><span data-stu-id="412ba-193">If your add-in is an Outlook add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }

    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

<span data-ttu-id="412ba-194">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-194">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a><span data-ttu-id="412ba-195">Modifications requises pour un add-in PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-195">Changes required for a PowerPoint add-in (JavaScript)</span></span>

<span data-ttu-id="412ba-196">Si votre add-in est un add-in PowerPoint créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="412ba-196">If your add-in is a PowerPoint add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="412ba-197">Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-197">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="412ba-198">Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-198">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="412ba-199">Recherchez `writeDataToPowerPoint` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-199">Find the `writeDataToPowerPoint` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. <span data-ttu-id="412ba-200">Supprimez la `writeDataToExcel` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-200">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="412ba-201">Supprimez la `writeDataToOutlook` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-201">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="412ba-202">Supprimez la `writeDataToWord` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-202">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="412ba-203">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-203">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a><span data-ttu-id="412ba-204">Modifications requises pour un add-in PowerPoint (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-204">Changes required for a PowerPoint add-in (TypeScript)</span></span>

<span data-ttu-id="412ba-205">Si votre add-in est un add-in PowerPoint qui a été créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts**, recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :</span><span class="sxs-lookup"><span data-stu-id="412ba-205">If your add-in is a PowerPoint add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

<span data-ttu-id="412ba-206">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-206">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-javascript"></a><span data-ttu-id="412ba-207">Modifications requises pour un add-in Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-207">Changes required for a Word add-in (JavaScript)</span></span>

<span data-ttu-id="412ba-208">Si votre add-in est un add-in Word créé avec JavaScript, a apporté les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="412ba-208">If your add-in is a Word add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="412ba-209">Recherchez `writeDataToOfficeDocument` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-209">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="412ba-210">Recherchez `filterUserProfileInfo` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-210">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="412ba-211">Recherchez `writeDataToWord` la fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="412ba-211">Find the `writeDataToWord` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="412ba-212">Supprimez la `writeDataToExcel` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-212">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="412ba-213">Supprimez la `writeDataToOutlook` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-213">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="412ba-214">Supprimez la `writeDataToPowerPoint` fonction.</span><span class="sxs-lookup"><span data-stu-id="412ba-214">Delete the `writeDataToPowerPoint` function.</span></span>

<span data-ttu-id="412ba-215">Une fois ces modifications apportées, passez directement à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-215">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-typescript"></a><span data-ttu-id="412ba-216">Modifications requises pour un add-in Word (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="412ba-216">Changes required for a Word add-in (TypeScript)</span></span>

<span data-ttu-id="412ba-217">Si votre add-in est un add-in Word créé avec TypeScript, ouvrez **./src/taskpane/taskpane.ts,** recherchez la fonction et remplacez-la par la fonction suivante `writeDataToOfficeDocument` :</span><span class="sxs-lookup"><span data-stu-id="412ba-217">If your add-in is a Word add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

<span data-ttu-id="412ba-218">Une fois ces modifications apportées, continuez à la [section](#try-it-out) Essayer de cet article pour tester votre add-in mis à jour.</span><span class="sxs-lookup"><span data-stu-id="412ba-218">After you've made these changes, continue to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="412ba-219">Try it out</span><span class="sxs-lookup"><span data-stu-id="412ba-219">Try it out</span></span>

<span data-ttu-id="412ba-220">Si votre compl?ment est un compl?ment Excel, Word ou PowerPoint, compl?ez les étapes de la section suivante pour l’essayer. Si votre compl?ment est un compl?ment Outlook, compl?ez les étapes dans la section [Outlook.](#outlook)</span><span class="sxs-lookup"><span data-stu-id="412ba-220">If your add-in is an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If your add-in is an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="412ba-221">Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="412ba-221">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="412ba-222">Pour tester un complément Excel, Word ou PowerPoint, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="412ba-222">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="412ba-223">Dans le dossier racine du projet, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et chargez une version test de votre application dans l’application cliente Office précédemment sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="412ba-223">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="412ba-224">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="412ba-224">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="412ba-225">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="412ba-225">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="412ba-226">Dans l’application cliente Office qui s’ouvre lorsque vous exécutez la commande précédente (c’est-à-dire, Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre [](sso-quickstart.md#configure-sso) de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’ouvrez-vous pour l’application.</span><span class="sxs-lookup"><span data-stu-id="412ba-226">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="412ba-227">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="412ba-227">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="412ba-228">Dans l’application client Office, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="412ba-228">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="412ba-229">L’image ci-après illustre ce bouton dans Excel.</span><span class="sxs-lookup"><span data-stu-id="412ba-229">The following image shows this button in Excel.</span></span>

    ![Screenshot showing highlighted add-in button in Excel ribbon](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="412ba-231">En bas du volet Des tâches, sélectionnez le bouton Lire **mon OneDrive** Entreprise pour lancer le processus d’pertinence.</span><span class="sxs-lookup"><span data-stu-id="412ba-231">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span>

5. <span data-ttu-id="412ba-232">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="412ba-232">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="412ba-233">Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Microsoft 365 (professionnel ou scolaire).</span><span class="sxs-lookup"><span data-stu-id="412ba-233">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="412ba-234">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="412ba-234">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Capture d’écran montrant la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="412ba-236">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="412ba-236">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="412ba-237">Le add-in lit les données du OneDrive Entreprise de l’utilisateur et écrit les noms des 10 principaux fichiers et dossiers dans le document.</span><span class="sxs-lookup"><span data-stu-id="412ba-237">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the document.</span></span> <span data-ttu-id="412ba-238">L’image suivante montre un exemple de noms de fichiers et de dossiers écrits dans une feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="412ba-238">The following image shows an example of file and folder names written to an Excel worksheet.</span></span>

    ![Capture d’écran montrant les informations OneDrive Entreprise dans la feuille de calcul Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="412ba-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="412ba-240">Outlook</span></span>

<span data-ttu-id="412ba-241">Pour tester un complément Outlook, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="412ba-241">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="412ba-242">Dans le dossier racine du projet, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et chargez une version test de votre application.</span><span class="sxs-lookup"><span data-stu-id="412ba-242">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="412ba-243">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="412ba-243">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="412ba-244">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="412ba-244">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="412ba-245">Il se peut également que vous deviez exécuter votre invite de commande ou votre terminal en tant qu'administrateur pour que les modifications soient effectuées.</span><span class="sxs-lookup"><span data-stu-id="412ba-245">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="412ba-246">Assurez-vous que vous êtes connecté à Outlook avec un utilisateur membre de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que celui que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’oD [SSO](sso-quickstart.md#configure-sso) pour l’application.</span><span class="sxs-lookup"><span data-stu-id="412ba-246">Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="412ba-247">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="412ba-247">Doing so establishes the appropriate conditions for SSO to succeed.</span></span>

3. <span data-ttu-id="412ba-248">Rédigez un nouveau message dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="412ba-248">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="412ba-249">Dans la fenêtre de composition du message, choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet du complément.</span><span class="sxs-lookup"><span data-stu-id="412ba-249">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran illustrant la fenêtre Outlook Composer un message et le bouton du ruban du complément mis en évidence](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="412ba-251">En bas du volet Des tâches, sélectionnez le bouton Lire **mon OneDrive** Entreprise pour lancer le processus d’pertinence.</span><span class="sxs-lookup"><span data-stu-id="412ba-251">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span>

6. <span data-ttu-id="412ba-252">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="412ba-252">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="412ba-253">Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Microsoft 365 (professionnel ou scolaire).</span><span class="sxs-lookup"><span data-stu-id="412ba-253">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="412ba-254">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="412ba-254">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Capture d’écran de la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="412ba-256">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="412ba-256">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="412ba-257">Le add-in lit les données du OneDrive Entreprise de l’utilisateur et écrit les noms des 10 principaux fichiers et dossiers dans le corps du message électronique.</span><span class="sxs-lookup"><span data-stu-id="412ba-257">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the body of the email message.</span></span>

    ![Capture d’écran montrant les informations OneDrive Entreprise dans la fenêtre composer un message Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="412ba-259">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="412ba-259">Next steps</span></span>

<span data-ttu-id="412ba-260">Félicitations, vous avez personnalisé avec succès la fonctionnalité du module de personnalisation de l’oDS que vous avez créée avec le générateur Yeoman dans le démarrage rapide de l’personnalisation [SSO.](sso-quickstart.md)</span><span class="sxs-lookup"><span data-stu-id="412ba-260">Congratulations, you've successfully customized the functionality of the SSO-enabled add-in that you created with the Yeoman generator in the [SSO quick start](sso-quickstart.md).</span></span> <span data-ttu-id="412ba-261">Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="412ba-261">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="412ba-262">Consultez aussi</span><span class="sxs-lookup"><span data-stu-id="412ba-262">See also</span></span>

- [<span data-ttu-id="412ba-263">Activer l’authentification unique pour des compléments Office</span><span class="sxs-lookup"><span data-stu-id="412ba-263">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="412ba-264">Démarrage rapide de l’authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="412ba-264">Single sign-on (SSO) quick start</span></span>](sso-quickstart.md)
- [<span data-ttu-id="412ba-265">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="412ba-265">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="412ba-266">Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="412ba-266">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)
