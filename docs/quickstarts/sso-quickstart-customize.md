---
title: Personnaliser votre complément compatible avec l’authentification unique Node.js
description: Découvrez comment personnaliser le complément à extension SSO que vous avez créé avec le générateur Yeoman.
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c1d292ed8ead40201dd035d6ae8e6997174ea477
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094483"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a><span data-ttu-id="201e0-103">Personnaliser votre complément compatible avec l’authentification unique Node.js</span><span class="sxs-lookup"><span data-stu-id="201e0-103">Customize your Node.js SSO-enabled add-in</span></span>

> [!IMPORTANT]
> <span data-ttu-id="201e0-104">Cet article s’appuie sur le complément à extension SSO créé en remplissant le démarrage rapide de l’authentification [unique (SSO)](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="201e0-104">This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md).</span></span> <span data-ttu-id="201e0-105">Veuillez terminer le démarrage rapide avant de lire cet article.</span><span class="sxs-lookup"><span data-stu-id="201e0-105">Please complete the quick start before reading this article.</span></span>

<span data-ttu-id="201e0-106">Le [démarrage rapide de l’authentification unique](sso-quickstart.md) crée un complément à extension SSO qui obtient les informations de profil de l’utilisateur connecté et l’écrit dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="201e0-106">The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document or message.</span></span> <span data-ttu-id="201e0-107">Dans cet article, vous découvrirez le processus de mise à jour du complément que vous avez créé avec le générateur Yeoman dans le démarrage rapide de l’authentification unique, afin d’ajouter de nouvelles fonctionnalités qui nécessitent des autorisations différentes.</span><span class="sxs-lookup"><span data-stu-id="201e0-107">In this article, you'll walk through the process of updating the add-in that you created with the Yeoman generator in the SSO quick start, to add new functionality that requires different permissions.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="201e0-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="201e0-108">Prerequisites</span></span>

* <span data-ttu-id="201e0-109">Un complément Office que vous avez créé en suivant les instructions du [démarrage rapide de l’authentification unique](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="201e0-109">An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).</span></span>

* <span data-ttu-id="201e0-110">Au moins quelques fichiers et dossiers stockés sur OneDrive entreprise dans votre abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="201e0-110">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="201e0-111">[Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="201e0-111">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a><span data-ttu-id="201e0-112">Vérifier le contenu du projet</span><span class="sxs-lookup"><span data-stu-id="201e0-112">Review contents of the project</span></span>

<span data-ttu-id="201e0-113">Commençons par un examen rapide du projet de complément que vous avez [créé précédemment avec le générateur Yeoman](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="201e0-113">Let's begin with a quick review of the add-in project that you previously [created with the Yeoman generator](sso-quickstart.md).</span></span>

> [!NOTE]
> <span data-ttu-id="201e0-114">Dans les emplacements où cet article fait référence à des fichiers de script utilisant l’extension de fichier **. js** , supposez plutôt l’extension de fichier **. TS** si votre projet a été créé avec une écriture.</span><span class="sxs-lookup"><span data-stu-id="201e0-114">In places where this article references script files using **.js** file extension, assume the **.ts** file extension instead if your project was created with TypeScript.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a><span data-ttu-id="201e0-115">Ajouter une nouvelle fonctionnalité</span><span class="sxs-lookup"><span data-stu-id="201e0-115">Add new functionality</span></span>

<span data-ttu-id="201e0-116">Le complément que vous avez créé avec le démarrage rapide de l’authentification unique utilise Microsoft Graph pour obtenir les informations de profil de l’utilisateur connecté et écrit ces informations dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="201e0-116">The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document or message.</span></span> <span data-ttu-id="201e0-117">Nous allons modifier les fonctionnalités du complément de sorte qu’il récupère les noms des 10 fichiers et dossiers les plus à partir de OneDrive entreprise de l’utilisateur connecté et qu’il écrit ces informations dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="201e0-117">Let's change the add-in's functionality such that it gets the names of the top 10 files and folders from the signed-in user's OneDrive for Business and writes that information to the document or message.</span></span> <span data-ttu-id="201e0-118">L’activation de cette nouvelle fonctionnalité nécessite la mise à jour des autorisations d’application dans Azure et la mise à jour du code dans le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="201e0-118">Enabling this new functionality requires updating app permissions in Azure and updating code within the add-in project.</span></span>

### <a name="update-app-permissions-in-azure"></a><span data-ttu-id="201e0-119">Mettre à jour les autorisations d’application dans Azure</span><span class="sxs-lookup"><span data-stu-id="201e0-119">Update app permissions in Azure</span></span>

<span data-ttu-id="201e0-120">Avant que le complément puisse lire correctement le contenu de OneDrive entreprise de l’utilisateur, les informations d’inscription de son application dans Azure doivent être mises à jour avec les autorisations appropriées.</span><span class="sxs-lookup"><span data-stu-id="201e0-120">Before the add-in can successfully read the contents of the user's OneDrive for Business, its app registration information in Azure must be updated with the appropriate permissions.</span></span> <span data-ttu-id="201e0-121">Procédez comme suit pour accorder à l’application l’autorisation **files. Read. All** et révoquer l’autorisation **User. Read** , qui n’est plus nécessaire.</span><span class="sxs-lookup"><span data-stu-id="201e0-121">Complete the following steps to grant the app the **Files.Read.All** permission and revoke the **User.Read** permission, which is no longer needed.</span></span>

1. <span data-ttu-id="201e0-122">Accédez au [portail Azure](https://ms.portal.azure.com/#home) et **Connectez-vous à l’aide de vos informations d’identification d’administrateur Microsoft 365**.</span><span class="sxs-lookup"><span data-stu-id="201e0-122">Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and **sign in using your Microsoft 365 administrator credentials**.</span></span>

2. <span data-ttu-id="201e0-123">Accédez à la page **inscriptions des applications** .</span><span class="sxs-lookup"><span data-stu-id="201e0-123">Navigate to the **App registrations** page.</span></span>
    > [!TIP]
    > <span data-ttu-id="201e0-124">Pour ce faire, vous pouvez choisir la vignette **inscriptions des applications** sur la page d’accueil Azure ou à l’aide de la zone de recherche de la page d’accueil pour rechercher et choisir les inscriptions de l' **application**.</span><span class="sxs-lookup"><span data-stu-id="201e0-124">You can do this either by choosing the **App registrations** tile on the Azure home page or by using the search box on the home page to find and choose **App registrations**.</span></span>

3. <span data-ttu-id="201e0-125">Sur la page **inscriptions des applications** , sélectionnez l’application que vous avez créée au démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="201e0-125">On the **App registrations** page, choose the app that you created during the quick start.</span></span> 
    > [!TIP]
    > <span data-ttu-id="201e0-126">Le **nom complet** de l’application correspond au nom du complément que vous avez spécifié lors de la création du projet avec le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="201e0-126">The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.</span></span>

4. <span data-ttu-id="201e0-127">À partir de la page vue d’ensemble de l’application, choisissez **autorisations d’API** sous l’en-tête **gérer** dans la partie gauche de la page.</span><span class="sxs-lookup"><span data-stu-id="201e0-127">From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.</span></span>

5. <span data-ttu-id="201e0-128">Dans la ligne **User. Read** du tableau autorisations, cliquez sur les points de suspension, puis sélectionnez **révoquer le consentement** de l’administrateur dans le menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="201e0-128">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Revoke admin consent** from the menu that appears.</span></span>

6. <span data-ttu-id="201e0-129">Sélectionnez le bouton **Oui, supprimer** en réponse à l’invite affichée.</span><span class="sxs-lookup"><span data-stu-id="201e0-129">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

7. <span data-ttu-id="201e0-130">Dans la ligne **User. Read** du tableau des autorisations, cliquez sur les points de suspension, puis sélectionnez **Supprimer l’autorisation** dans le menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="201e0-130">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Remove permission** from the menu that appears.</span></span>

8. <span data-ttu-id="201e0-131">Sélectionnez le bouton **Oui, supprimer** en réponse à l’invite affichée.</span><span class="sxs-lookup"><span data-stu-id="201e0-131">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

9. <span data-ttu-id="201e0-132">Cliquez sur le bouton **Ajouter une autorisation**.</span><span class="sxs-lookup"><span data-stu-id="201e0-132">Select the **Add a permission** button.</span></span>

10. <span data-ttu-id="201e0-133">Dans le panneau qui s’ouvre, choisissez **Microsoft Graph** , puis sélectionnez **autorisations déléguées**.</span><span class="sxs-lookup"><span data-stu-id="201e0-133">On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

11. <span data-ttu-id="201e0-134">Dans le panneau autorisations de l' **API de demande** :</span><span class="sxs-lookup"><span data-stu-id="201e0-134">On the **Request API permissions** panel:</span></span>

    <span data-ttu-id="201e0-135">a.</span><span class="sxs-lookup"><span data-stu-id="201e0-135">a.</span></span> <span data-ttu-id="201e0-136">Sous **fichiers**, sélectionnez **fichiers. Read. All**.</span><span class="sxs-lookup"><span data-stu-id="201e0-136">Under **Files**, select **Files.Read.All**.</span></span>

    <span data-ttu-id="201e0-137">b.</span><span class="sxs-lookup"><span data-stu-id="201e0-137">b.</span></span> <span data-ttu-id="201e0-138">Sélectionnez le bouton **Ajouter des autorisations** en bas du panneau pour enregistrer ces modifications d’autorisations.</span><span class="sxs-lookup"><span data-stu-id="201e0-138">Select the **Add permissions** button at the bottom of the panel to save these permissions changes.</span></span>

12. <span data-ttu-id="201e0-139">Sélectionnez le bouton **accorder le consentement de l’administrateur pour [nom du client]** .</span><span class="sxs-lookup"><span data-stu-id="201e0-139">Select the **Grant admin consent for [tenant name]** button.</span></span>

13. <span data-ttu-id="201e0-140">Sélectionnez le bouton **Oui** en réponse à l’invite affichée.</span><span class="sxs-lookup"><span data-stu-id="201e0-140">Select the **Yes** button in response to the prompt that's displayed.</span></span>

### <a name="update-code-in-the-add-in-project"></a><span data-ttu-id="201e0-141">Mettre à jour le code dans le projet de complément</span><span class="sxs-lookup"><span data-stu-id="201e0-141">Update code in the add-in project</span></span>

<span data-ttu-id="201e0-142">Pour permettre au complément de lire le contenu de OneDrive entreprise de l’utilisateur connecté, vous devez :</span><span class="sxs-lookup"><span data-stu-id="201e0-142">To enable the add-in to read contents of the signed-in user's OneDrive for Business, you'll need to:</span></span>

- <span data-ttu-id="201e0-143">Mettez à jour le code qui fait référence à l’URL de Microsoft Graph, aux paramètres et à l’étendue d’accès requise.</span><span class="sxs-lookup"><span data-stu-id="201e0-143">Update the code that references the Microsoft Graph URL, parameters, and required access scope.</span></span>

- <span data-ttu-id="201e0-144">Mettez à jour le code qui définit l’interface utilisateur du volet Office, afin qu’il décrive précisément les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="201e0-144">Update the code that defines the task pane UI, so that it accurately describes the new functionality.</span></span> 

- <span data-ttu-id="201e0-145">Mettez à jour le code qui analyse la réponse à partir de Microsoft Graph et l’écrit dans le document ou le message.</span><span class="sxs-lookup"><span data-stu-id="201e0-145">Update the code that parses the response from Microsoft Graph and writes it to the document or message.</span></span>

<span data-ttu-id="201e0-146">Les étapes suivantes décrivent ces mises à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-146">The following steps describe these updates.</span></span>

### <a name="changes-required-for-any-type-of-add-in"></a><span data-ttu-id="201e0-147">Modifications requises pour tout type de complément</span><span class="sxs-lookup"><span data-stu-id="201e0-147">Changes required for any type of add-in</span></span>

<span data-ttu-id="201e0-148">Effectuez les étapes suivantes pour votre complément, pour modifier l’URL, les paramètres et l’étendue d’accès de Microsoft Graph, et mettre à jour l’interface utilisateur du volet Office.</span><span class="sxs-lookup"><span data-stu-id="201e0-148">Complete the following steps for your add-in, to change the Microsoft Graph URL, parameters, and access scope, and update the taskpane UI.</span></span> <span data-ttu-id="201e0-149">Ces étapes sont les mêmes, quel que soit l’hôte Office que votre complément cible.</span><span class="sxs-lookup"><span data-stu-id="201e0-149">These steps are the same, regardless of which Office host your add-in targets.</span></span>

1. <span data-ttu-id="201e0-150">Dans le **./. ENV** (fichier) :</span><span class="sxs-lookup"><span data-stu-id="201e0-150">In the **./.ENV** file:</span></span>

    <span data-ttu-id="201e0-151">a.</span><span class="sxs-lookup"><span data-stu-id="201e0-151">a.</span></span> <span data-ttu-id="201e0-152">Remplacez `GRAPH_URL_SEGMENT=/me` par :`GRAPH_URL_SEGMENT=/me/drive/root/children`</span><span class="sxs-lookup"><span data-stu-id="201e0-152">Replace `GRAPH_URL_SEGMENT=/me` with the following: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span></span>

    <span data-ttu-id="201e0-153">b.</span><span class="sxs-lookup"><span data-stu-id="201e0-153">b.</span></span> <span data-ttu-id="201e0-154">Remplacez `QUERY_PARAM_SEGMENT=` par :`QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span><span class="sxs-lookup"><span data-stu-id="201e0-154">Replace `QUERY_PARAM_SEGMENT=` with the following: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span></span>

    <span data-ttu-id="201e0-155">c.</span><span class="sxs-lookup"><span data-stu-id="201e0-155">c.</span></span> <span data-ttu-id="201e0-156">Remplacez `SCOPE=User.Read` par :`SCOPE=Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="201e0-156">Replace `SCOPE=User.Read` with the following: `SCOPE=Files.Read.All`</span></span>

2. <span data-ttu-id="201e0-157">Dans **./manifest.xml**, recherchez la ligne à `<Scope>User.Read</Scope>` la fin du fichier et remplacez-la par la ligne `<Scope>Files.Read.All</Scope>` .</span><span class="sxs-lookup"><span data-stu-id="201e0-157">In **./manifest.xml**, find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>`.</span></span>

3. <span data-ttu-id="201e0-158">Dans **./src/helpers/fallbackauthdialog.js** (ou dans **./SRC/helpers/fallbackauthdialog.TS** pour un projet de type dactylographié), recherchez la chaîne `https://graph.microsoft.com/User.Read` et remplacez-la par la chaîne, de la manière suivante `https://graph.microsoft.com/Files.Read.All` `requestObj` :</span><span class="sxs-lookup"><span data-stu-id="201e0-158">In **./src/helpers/fallbackauthdialog.js** (or in **./src/helpers/fallbackauthdialog.ts** for a TypeScript project), find the string `https://graph.microsoft.com/User.Read` and replace it with the string `https://graph.microsoft.com/Files.Read.All`, such that `requestObj` is defined as follows:</span></span>

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

4. <span data-ttu-id="201e0-159">Dans **./src/taskpane/taskpane.html**, recherchez l’élément `<section class="ms-firstrun-instructionstep__header">` et mettez à jour le texte à l’intérieur de cet élément pour décrire les nouvelles fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="201e0-159">In **./src/taskpane/taskpane.html**, find the element `<section class="ms-firstrun-instructionstep__header">` and update the text within that element to describe the add-in's new functionality.</span></span>

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. <span data-ttu-id="201e0-160">Dans **./src/taskpane/taskpane.html**, recherchez et remplacez les deux occurrences de la chaîne `Get My User Profile Information` par la chaîne `Read my OneDrive for Business` .</span><span class="sxs-lookup"><span data-stu-id="201e0-160">In **./src/taskpane/taskpane.html**, find and replace both occurrences of the string `Get My User Profile Information` with the string `Read my OneDrive for Business`.</span></span>

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

6. <span data-ttu-id="201e0-161">Dans **./src/taskpane/taskpane.html**, recherchez et remplacez la chaîne `Your user profile information will be displayed in the document.` par la chaîne `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .</span><span class="sxs-lookup"><span data-stu-id="201e0-161">In **./src/taskpane/taskpane.html**, find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. <span data-ttu-id="201e0-162">Mettez à jour le code qui analyse la réponse à partir de Microsoft Graph et l’écrit dans le document ou le message, en suivant les instructions de la section correspondant à votre type de complément :</span><span class="sxs-lookup"><span data-stu-id="201e0-162">Update the code that parses the response from Microsoft Graph and writes it to the document or message, by following guidance in the section that corresponds to your type of add-in:</span></span>

    - [<span data-ttu-id="201e0-163">Modifications requises pour un complément Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-163">Changes required for an Excel add-in (JavaScript)</span></span>](#changes-required-for-an-excel-add-in-javascript)
    - [<span data-ttu-id="201e0-164">Modifications requises pour un complément Excel (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-164">Changes required for an Excel add-in (TypeScript)</span></span>](#changes-required-for-an-excel-add-in-typescript)
    - [<span data-ttu-id="201e0-165">Modifications requises pour un complément Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-165">Changes required for an Outlook add-in (JavaScript)</span></span>](#changes-required-for-an-outlook-add-in-javascript)
    - [<span data-ttu-id="201e0-166">Modifications requises pour un complément Outlook (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-166">Changes required for an Outlook add-in (TypeScript)</span></span>](#changes-required-for-an-outlook-add-in-typescript)
    - [<span data-ttu-id="201e0-167">Modifications requises pour un complément PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-167">Changes required for a PowerPoint add-in (JavaScript)</span></span>](#changes-required-for-a-powerpoint-add-in-javascript)
    - [<span data-ttu-id="201e0-168">Modifications requises pour un complément PowerPoint (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-168">Changes required for a PowerPoint add-in (TypeScript)</span></span>](#changes-required-for-a-powerpoint-add-in-typescript)
    - [<span data-ttu-id="201e0-169">Modifications requises pour un complément Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-169">Changes required for a Word add-in (JavaScript)</span></span>](#changes-required-for-a-word-add-in-javascript)
    - [<span data-ttu-id="201e0-170">Modifications requises pour un complément Word (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-170">Changes required for a Word add-in (TypeScript)</span></span>](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a><span data-ttu-id="201e0-171">Modifications requises pour un complément Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-171">Changes required for an Excel add-in (JavaScript)</span></span>

<span data-ttu-id="201e0-172">Si votre complément est un complément Excel qui a été créé avec JavaScript, effectuez les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="201e0-172">If your add-in is an Excel add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="201e0-173">Recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-173">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="201e0-174">Recherchez la `filterUserProfileInfo` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-174">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="201e0-175">Recherchez la `writeDataToExcel` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-175">Find the `writeDataToExcel` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="201e0-176">Supprimez la `writeDataToOutlook` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-176">Delete the `writeDataToOutlook` function.</span></span>

5. <span data-ttu-id="201e0-177">Supprimez la `writeDataToPowerPoint` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-177">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="201e0-178">Supprimez la `writeDataToWord` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-178">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="201e0-179">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-179">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-excel-add-in-typescript"></a><span data-ttu-id="201e0-180">Modifications requises pour un complément Excel (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-180">Changes required for an Excel add-in (TypeScript)</span></span>

<span data-ttu-id="201e0-181">Si votre complément est un complément Excel qui a été créé avec la machine à écrire, ouvrez **./SRC/TaskPane/TaskPane.TS**, recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-181">If your add-in is an Excel add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="201e0-182">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-182">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-javascript"></a><span data-ttu-id="201e0-183">Modifications requises pour un complément Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-183">Changes required for an Outlook add-in (JavaScript)</span></span>

<span data-ttu-id="201e0-184">Si votre complément est un complément Outlook créé avec JavaScript, effectuez les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="201e0-184">If your add-in is an Outlook add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="201e0-185">Recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-185">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="201e0-186">Recherchez la `filterUserProfileInfo` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-186">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="201e0-187">Recherchez la `writeDataToOutlook` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-187">Find the `writeDataToOutlook` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="201e0-188">Supprimez la `writeDataToExcel` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-188">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="201e0-189">Supprimez la `writeDataToPowerPoint` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-189">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="201e0-190">Supprimez la `writeDataToWord` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-190">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="201e0-191">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-191">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-typescript"></a><span data-ttu-id="201e0-192">Modifications requises pour un complément Outlook (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-192">Changes required for an Outlook add-in (TypeScript)</span></span>

<span data-ttu-id="201e0-193">Si votre complément est un complément Outlook créé avec la machine à écrire, ouvrez **./SRC/TaskPane/TaskPane.TS**, recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-193">If your add-in is an Outlook add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="201e0-194">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-194">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a><span data-ttu-id="201e0-195">Modifications requises pour un complément PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-195">Changes required for a PowerPoint add-in (JavaScript)</span></span>

<span data-ttu-id="201e0-196">Si votre complément est un complément PowerPoint créé avec JavaScript, effectuez les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="201e0-196">If your add-in is a PowerPoint add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="201e0-197">Recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-197">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="201e0-198">Recherchez la `filterUserProfileInfo` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-198">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="201e0-199">Recherchez la `writeDataToPowerPoint` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-199">Find the `writeDataToPowerPoint` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="201e0-200">Supprimez la `writeDataToExcel` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-200">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="201e0-201">Supprimez la `writeDataToOutlook` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-201">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="201e0-202">Supprimez la `writeDataToWord` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-202">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="201e0-203">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-203">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a><span data-ttu-id="201e0-204">Modifications requises pour un complément PowerPoint (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-204">Changes required for a PowerPoint add-in (TypeScript)</span></span>

<span data-ttu-id="201e0-205">Si votre complément est un complément PowerPoint créé avec la machine à écrire, ouvrez **./SRC/TaskPane/TaskPane.TS**, recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-205">If your add-in is a PowerPoint add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="201e0-206">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-206">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-javascript"></a><span data-ttu-id="201e0-207">Modifications requises pour un complément Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="201e0-207">Changes required for a Word add-in (JavaScript)</span></span>

<span data-ttu-id="201e0-208">Si votre complément est un complément Word créé avec JavaScript, effectuez les modifications suivantes dans **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="201e0-208">If your add-in is a Word add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="201e0-209">Recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-209">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="201e0-210">Recherchez la `filterUserProfileInfo` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-210">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="201e0-211">Recherchez la `writeDataToWord` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-211">Find the `writeDataToWord` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="201e0-212">Supprimez la `writeDataToExcel` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-212">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="201e0-213">Supprimez la `writeDataToOutlook` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-213">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="201e0-214">Supprimez la `writeDataToPowerPoint` fonction.</span><span class="sxs-lookup"><span data-stu-id="201e0-214">Delete the `writeDataToPowerPoint` function.</span></span>

<span data-ttu-id="201e0-215">Une fois ces modifications effectuées, passez directement à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-215">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-typescript"></a><span data-ttu-id="201e0-216">Modifications requises pour un complément Word (machine à écrire)</span><span class="sxs-lookup"><span data-stu-id="201e0-216">Changes required for a Word add-in (TypeScript)</span></span>

<span data-ttu-id="201e0-217">Si votre complément est un complément Word qui a été créé avec une machine à écrire, ouvrez **./SRC/TaskPane/TaskPane.TS**, recherchez la `writeDataToOfficeDocument` fonction et remplacez-la par la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="201e0-217">If your add-in is a Word add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="201e0-218">Une fois ces modifications effectuées, passez à la section [essayer](#try-it-out) de cet article pour tester votre complément mis à jour.</span><span class="sxs-lookup"><span data-stu-id="201e0-218">After you've made these changes, continue to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="201e0-219">Try it out</span><span class="sxs-lookup"><span data-stu-id="201e0-219">Try it out</span></span>

<span data-ttu-id="201e0-220">Si votre complément est un complément Excel, Word ou PowerPoint, effectuez les étapes de la section suivante pour le tester. Si votre complément est un complément Outlook, effectuez plutôt les étapes dans la section [Outlook](#outlook) .</span><span class="sxs-lookup"><span data-stu-id="201e0-220">If your add-in is an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If your add-in is an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="201e0-221">Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="201e0-221">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="201e0-222">Pour tester un complément Excel, Word ou PowerPoint, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="201e0-222">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="201e0-223">Dans le dossier racine du projet, exécutez la commande suivante pour générer le projet, démarrez le serveur Web local et chargement votre complément dans l’application cliente Office précédemment sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="201e0-223">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="201e0-224">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="201e0-224">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="201e0-225">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="201e0-225">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="201e0-226">Dans l’application cliente Office qui s’ouvre lorsque vous exécutez la commande précédente (Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre de la même organisation 365 Microsoft que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l' [authentification unique](sso-quickstart.md#configure-sso) pour l’application.</span><span class="sxs-lookup"><span data-stu-id="201e0-226">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="201e0-227">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="201e0-227">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="201e0-228">Dans l’application client Office, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="201e0-228">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="201e0-229">L’image ci-après illustre ce bouton dans Excel.</span><span class="sxs-lookup"><span data-stu-id="201e0-229">The following image shows this button in Excel.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="201e0-231">En bas du volet Office, cliquez sur le bouton **lire mon OneDrive entreprise** pour lancer le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="201e0-231">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="201e0-232">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="201e0-232">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="201e0-233">Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph ou lorsque l’utilisateur n’est pas connecté à Office avec un compte Microsoft valide ou un compte professionnel ou scolaire Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="201e0-233">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="201e0-234">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="201e0-234">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="201e0-236">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="201e0-236">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="201e0-237">Le complément lit les données de OneDrive entreprise de l’utilisateur connecté et écrit les noms des 10 premiers fichiers et dossiers dans le document.</span><span class="sxs-lookup"><span data-stu-id="201e0-237">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the document.</span></span> <span data-ttu-id="201e0-238">L’image suivante montre un exemple de noms de fichiers et de dossiers écrits dans une feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="201e0-238">The following image shows an example of file and folder names written to an Excel worksheet.</span></span>

    ![Informations OneDrive entreprise dans la feuille de calcul Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="201e0-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="201e0-240">Outlook</span></span>

<span data-ttu-id="201e0-241">Pour tester un complément Outlook, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="201e0-241">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="201e0-242">Dans le dossier racine du projet, exécutez la commande suivante pour générer le projet et démarrer le serveur Web local.</span><span class="sxs-lookup"><span data-stu-id="201e0-242">In the root folder of the project, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="201e0-243">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="201e0-243">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="201e0-244">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="201e0-244">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="201e0-245">Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing) pour charger le complément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="201e0-245">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="201e0-246">Assurez-vous que vous êtes connecté à Outlook avec un utilisateur membre de la même organisation 365 Microsoft que le compte administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l' [authentification unique](sso-quickstart.md#configure-sso) pour l’application.</span><span class="sxs-lookup"><span data-stu-id="201e0-246">Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="201e0-247">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="201e0-247">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="201e0-248">Rédigez un nouveau message dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="201e0-248">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="201e0-249">Dans la fenêtre de composition du message, choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet du complément.</span><span class="sxs-lookup"><span data-stu-id="201e0-249">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton du complément Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="201e0-251">En bas du volet Office, cliquez sur le bouton **lire mon OneDrive entreprise** pour lancer le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="201e0-251">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="201e0-252">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="201e0-252">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="201e0-253">Cela peut se produire lorsque l’administrateur client n’a pas accordé le consentement du complément pour accéder à Microsoft Graph ou lorsque l’utilisateur n’est pas connecté à Office avec un compte Microsoft valide ou un compte professionnel ou scolaire Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="201e0-253">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="201e0-254">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="201e0-254">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="201e0-256">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="201e0-256">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="201e0-257">Le complément lit les données de OneDrive entreprise de l’utilisateur connecté et écrit les noms des 10 fichiers et dossiers les plus fréquents dans le corps du message électronique.</span><span class="sxs-lookup"><span data-stu-id="201e0-257">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the body of the email message.</span></span>

    ![Informations OneDrive entreprise dans un message Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="201e0-259">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="201e0-259">Next steps</span></span>

<span data-ttu-id="201e0-260">Félicitations, vous avez personnalisé avec succès les fonctionnalités du complément à extension SSO que vous avez créé avec le générateur Yeoman dans le [démarrage rapide de l’authentification unique](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="201e0-260">Congratulations, you've successfully customized the functionality of the SSO-enabled add-in that you created with the Yeoman generator in the [SSO quick start](sso-quickstart.md).</span></span> <span data-ttu-id="201e0-261">Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="201e0-261">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="201e0-262">Consultez aussi</span><span class="sxs-lookup"><span data-stu-id="201e0-262">See also</span></span>

- [<span data-ttu-id="201e0-263">Activer l’authentification unique pour des compléments Office</span><span class="sxs-lookup"><span data-stu-id="201e0-263">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="201e0-264">Démarrage rapide de l’authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="201e0-264">Single sign-on (SSO) quick start</span></span>](sso-quickstart.md)
- [<span data-ttu-id="201e0-265">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="201e0-265">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="201e0-266">Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="201e0-266">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)
