---
ms.date: 06/17/2019
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f172139f3aafb374eec3c1350b127ed3194d00e0
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059929"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="e3e9c-103">Prise en main du développement de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="e3e9c-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="e3e9c-104">Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions à Excel en les définissant en JavaScript ou en une machine à écrire dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="e3e9c-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe `SUM()`quelle fonction native dans Excel, comme.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e3e9c-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="e3e9c-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="e3e9c-107">Excel sur Windows (version 1904 ou ultérieure, connexion à l’abonnement Office 365) ou Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="e3e9c-107">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="e3e9c-108">Les fonctions personnalisées d’Excel sont prises en charge dans Office sur Mac (connexion à l’abonnement Office 365) et une mise à jour de ce didacticiel est prochainement prévue.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-108">Excel custom functions are supported in Office on Mac (connected to Office 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="e3e9c-109">Les fonctions personnalisées d’Excel ne sont pas prises en charge dans Office 2019 (achat unique).</span><span class="sxs-lookup"><span data-stu-id="e3e9c-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="e3e9c-110">Création de votre premier projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e3e9c-110">Build your first custom functions project</span></span>

<span data-ttu-id="e3e9c-111">Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="e3e9c-112">Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="e3e9c-113">Dans un dossier de votre choix, exécutez la commande suivante, puis répondez aux invites comme suit.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-113">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="e3e9c-114">**Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="e3e9c-114">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="e3e9c-115">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="e3e9c-115">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="e3e9c-116">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="e3e9c-116">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/UpdatedYoOfficePrompt.png)

    <span data-ttu-id="e3e9c-118">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-118">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="e3e9c-119">Le générateur Yeoman vous donne des instructions dans votre ligne de commande sur ce qu’il faut faire du projet, mais il les ignore et continue de suivre nos instructions.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-119">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="e3e9c-120">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="e3e9c-121">Créez le projet.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-121">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="e3e9c-122">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="e3e9c-123">Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-123">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="e3e9c-124">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-124">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="e3e9c-125">Vous pouvez essayer le complément de fonction personnalisée dans Excel sur Windows ou Excel online.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-125">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span> <span data-ttu-id="e3e9c-126">Vous serez peut-être invité à ouvrir le volet Office du complément, bien que ce soit facultatif.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-126">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="e3e9c-127">Vous pouvez toujours exécuter vos fonctions personnalisées sans ouvrir le volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-127">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="e3e9c-128">Excel sur Windows</span><span class="sxs-lookup"><span data-stu-id="e3e9c-128">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="e3e9c-129">Pour tester votre complément dans Excel sous Windows, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-129">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="e3e9c-130">Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="e3e9c-131">Excel Online</span><span class="sxs-lookup"><span data-stu-id="e3e9c-131">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="e3e9c-132">Pour tester votre complément dans Excel Online, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-132">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="e3e9c-133">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="e3e9c-134">Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel online.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-134">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="e3e9c-135">Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-135">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="e3e9c-136">Dans Excel Online, sélectionnez l’onglet **Insérer**, puis **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-136">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Insérer un ruban dans Excel Online avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="e3e9c-138">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-138">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="e3e9c-139">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-139">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="e3e9c-140">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-140">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="e3e9c-141">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="e3e9c-141">Try out a prebuilt custom function</span></span>

<span data-ttu-id="e3e9c-142">Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Yeoman contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="e3e9c-142">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="e3e9c-143">Le fichier **./manifest.xml** dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à `CONTOSO` l’espace de noms.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-143">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="e3e9c-144">Dans votre classeur Excel, essayez la `ADD` fonction personnalisée en procédant comme suit:</span><span class="sxs-lookup"><span data-stu-id="e3e9c-144">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="e3e9c-145">Sélectionnez une cellule et tapez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-145">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="e3e9c-146">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-146">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="e3e9c-147">Exécutez la `CONTOSO.ADD` fonction, en utilisant `10` des `200` nombres et comme paramètres d’entrée, en `=CONTOSO.ADD(10,200)` tapant la valeur dans la cellule et en appuyant sur entrée.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-147">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="e3e9c-148">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-148">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="e3e9c-149">La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-149">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e3e9c-150">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="e3e9c-150">Next steps</span></span>

<span data-ttu-id="e3e9c-151">Félicitations, vous avez créé une fonction personnalisée dans un complément Excel!</span><span class="sxs-lookup"><span data-stu-id="e3e9c-151">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="e3e9c-152">Ensuite, créez un complément plus complexe avec la fonctionnalité de diffusion de données en continu.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-152">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="e3e9c-153">Le lien suivant vous guide tout au long des étapes suivantes du didacticiel de complément Excel avec fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="e3e9c-153">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e3e9c-154">Didacticiel de complément de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="e3e9c-154">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="e3e9c-155">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e3e9c-155">See also</span></span>

* [<span data-ttu-id="e3e9c-156">Vue d’ensemble des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e3e9c-156">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="e3e9c-157">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e3e9c-157">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="e3e9c-158">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="e3e9c-158">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="e3e9c-159">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e3e9c-159">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
