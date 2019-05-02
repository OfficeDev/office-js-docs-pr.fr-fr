---
ms.date: 03/06/2019
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées (aperçu)
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 3ea7ec4c2089aaa4e9f193a45e7c4a31c691f213
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517071"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="48a4b-103">Prise en main du développement de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="48a4b-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="48a4b-104">Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions à Excel en les définissant en JavaScript ou en une machine à écrire dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="48a4b-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="48a4b-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe `SUM()`quelle fonction native dans Excel, comme.</span><span class="sxs-lookup"><span data-stu-id="48a4b-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="48a4b-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="48a4b-106">Prerequisites</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="48a4b-107">Vous aurez besoin des outils et ressources connexes suivants pour commencer à créer des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="48a4b-107">You'll need the following tools and related resources to begin creating custom functions.</span></span>

- <span data-ttu-id="48a4b-108">[Node.js](https://nodejs.org/en/) (version 8.0.0 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="48a4b-108">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="48a4b-109">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="48a4b-109">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="48a4b-110">La dernière version de[Yeoman](https://yeoman.io/) et de [Yeoman Générateur de compléments Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="48a4b-110">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="48a4b-111">Même si vous avez déjà installé le générateur Yeoman, nous vous recommandons de mettre à jour votre package vers la dernière version à partir de NPM.</span><span class="sxs-lookup"><span data-stu-id="48a4b-111">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="48a4b-112">Création de votre premier projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="48a4b-112">Build your first custom functions project</span></span>

<span data-ttu-id="48a4b-113">Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="48a4b-113">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="48a4b-114">Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="48a4b-114">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="48a4b-115">Exécutez la commande suivante, puis répondez aux invitations comme suit.</span><span class="sxs-lookup"><span data-stu-id="48a4b-115">Run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="48a4b-116">Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="48a4b-116">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    - <span data-ttu-id="48a4b-117">Choisissez un type de script : `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="48a4b-117">Choose a script type: `JavaScript`</span></span>

    - <span data-ttu-id="48a4b-118">Comment souhaitez-vous nommer votre complément ?</span><span class="sxs-lookup"><span data-stu-id="48a4b-118">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="48a4b-120">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="48a4b-120">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="48a4b-121">Naviguez jusqu’au dossier de projet que vous venez de créer.</span><span class="sxs-lookup"><span data-stu-id="48a4b-121">Navigate to the project folder you just created.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="48a4b-122">Approuvez le certificat auto-signé dont vous avez besoin pour exécuter ce projet.</span><span class="sxs-lookup"><span data-stu-id="48a4b-122">Trust the self-signed certificate you need to run this project.</span></span> <span data-ttu-id="48a4b-123">Pour obtenir des instructions détaillées pour Windows ou Mac, voir [Ajout des Certificats Auto-signés comme Certificat Racine Approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="48a4b-123">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="48a4b-124">Construire le projet.</span><span class="sxs-lookup"><span data-stu-id="48a4b-124">Build the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

5. <span data-ttu-id="48a4b-125">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="48a4b-125">Start the local web server, which runs in Node.js.</span></span>

    - <span data-ttu-id="48a4b-126">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur Web local, lancez Excel et chargement le complément:</span><span class="sxs-lookup"><span data-stu-id="48a4b-126">If you use Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```command&nbsp;line
         npm run start
        ```
        <span data-ttu-id="48a4b-127">Après avoir exécuté cette commande, votre invite de commandes affiche des détails sur le démarrage du serveur Web.</span><span class="sxs-lookup"><span data-stu-id="48a4b-127">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="48a4b-128">Excel commence avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="48a4b-128">Excel will start with your add-in loaded.</span></span> <span data-ttu-id="48a4b-129">Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3.</span><span class="sxs-lookup"><span data-stu-id="48a4b-129">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    - <span data-ttu-id="48a4b-130">Si vous utilisez Excel Online pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur Web local:</span><span class="sxs-lookup"><span data-stu-id="48a4b-130">If you use Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```command&nbsp;line
        npm run start-web
        ```

         <span data-ttu-id="48a4b-131">Après avoir exécuté cette commande, votre invite de commandes affiche des détails sur le démarrage du serveur Web.</span><span class="sxs-lookup"><span data-stu-id="48a4b-131">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="48a4b-132">Pour utiliser vos fonctions, ouvrez un nouveau classeur dans Excel online.</span><span class="sxs-lookup"><span data-stu-id="48a4b-132">To use your functions, open a new workbook in Excel Online.</span></span> <span data-ttu-id="48a4b-133">Dans ce classeur, vous devrez charger votre complément.</span><span class="sxs-lookup"><span data-stu-id="48a4b-133">In this workbook, you'll need to load your add-in.</span></span> 

        <span data-ttu-id="48a4b-134">Pour ce faire, sélectionnez l’onglet **Insérer** sur le ruban et sélectionnez **Get Add-ins**. Dans la nouvelle fenêtre qui s’affiche, vérifiez que vous êtes dans l’onglet **mes compléments** . Ensuite, sélectionnez **gérer mes compléments _GT_ Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="48a4b-134">To do this, select the **Insert** tab on the ribbon and select **Get Add-ins**. In the resulting new window, ensure you are on the **My Add-ins** tab. Next, select **Manage My Add-ins > Upload My Add-in**.</span></span> <span data-ttu-id="48a4b-135">Recherchez votre fichier manifeste et téléchargez-le.</span><span class="sxs-lookup"><span data-stu-id="48a4b-135">Browse for your manifest file and upload it.</span></span> <span data-ttu-id="48a4b-136">Si votre complément ne se charge pas, vérifiez que vous avez correctement terminé l’étape 3.</span><span class="sxs-lookup"><span data-stu-id="48a4b-136">If your add-in does not load, check you've completed step 3 correctly.</span></span>

## <a name="try-out-the-prebuilt-custom-functions"></a><span data-ttu-id="48a4b-137">Tester les fonctions personnalisées prédéfinies</span><span class="sxs-lookup"><span data-stu-id="48a4b-137">Try out the prebuilt custom functions</span></span>

<span data-ttu-id="48a4b-138">Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Office Yo contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="48a4b-138">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="48a4b-139">Le fichier**manifest.xml**dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à l’ `CONTOSO` espace de noms.</span><span class="sxs-lookup"><span data-stu-id="48a4b-139">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="48a4b-140">Dans votre classeur Excel, essayez la `ADD` fonction personnalisée en procédant comme suit:</span><span class="sxs-lookup"><span data-stu-id="48a4b-140">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="48a4b-141">Sélectionnez une cellule et tapez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="48a4b-141">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="48a4b-142">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="48a4b-142">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="48a4b-143">Exécutez la `CONTOSO.ADD` fonction, en utilisant `10` des `200` nombres et comme paramètres d’entrée, en `=CONTOSO.ADD(10,200)` tapant la valeur dans la cellule et en appuyant sur entrée.</span><span class="sxs-lookup"><span data-stu-id="48a4b-143">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="48a4b-144">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="48a4b-144">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="48a4b-145">La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="48a4b-145">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="48a4b-146">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="48a4b-146">Next steps</span></span>

<span data-ttu-id="48a4b-147">Félicitations, vous avez créé une fonction personnalisée dans un complément Excel!</span><span class="sxs-lookup"><span data-stu-id="48a4b-147">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="48a4b-148">Ensuite, créez un complément plus complexe avec la fonctionnalité de diffusion de données en continu.</span><span class="sxs-lookup"><span data-stu-id="48a4b-148">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="48a4b-149">Le lien suivant vous guide tout au long des étapes suivantes du didacticiel de complément Excel avec fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="48a4b-149">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="48a4b-150">Didacticiel de complément de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="48a4b-150">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="48a4b-151">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="48a4b-151">See also</span></span>

* [<span data-ttu-id="48a4b-152">Vue d’ensemble des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="48a4b-152">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="48a4b-153">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="48a4b-153">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="48a4b-154">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="48a4b-154">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="48a4b-155">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="48a4b-155">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
