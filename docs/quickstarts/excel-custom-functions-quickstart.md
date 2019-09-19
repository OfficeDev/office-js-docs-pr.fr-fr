---
ms.date: 09/18/2019
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f34a8817a7c8ef2679fc8ce0a6ad17cec600531b
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035328"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="0887f-103">Prise en main du développement de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="0887f-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="0887f-104">Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions à Excel en les définissant en JavaScript ou en une machine à écrire dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="0887f-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="0887f-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe `SUM()`quelle fonction native dans Excel, comme.</span><span class="sxs-lookup"><span data-stu-id="0887f-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0887f-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="0887f-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="0887f-107">Excel sur Windows (version 1904 ou ultérieure, connexion à l’abonnement Office 365) ou Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="0887f-107">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="0887f-108">Les fonctions personnalisées d’Excel sont prises en charge dans Office sur Mac (connexion à l’abonnement Office 365) et une mise à jour de ce didacticiel est prochainement prévue.</span><span class="sxs-lookup"><span data-stu-id="0887f-108">Excel custom functions are supported in Office on Mac (connected to Office 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="0887f-109">Les fonctions personnalisées d’Excel ne sont pas prises en charge dans Office 2019 (achat unique).</span><span class="sxs-lookup"><span data-stu-id="0887f-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="0887f-110">Création de votre premier projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="0887f-110">Build your first custom functions project</span></span>

<span data-ttu-id="0887f-111">Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0887f-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="0887f-112">Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0887f-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="0887f-113">**Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="0887f-113">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="0887f-114">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="0887f-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="0887f-115">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="0887f-115">**What do you want to name your add-in?**</span></span> `starcount`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/starcountPrompt.png)

    <span data-ttu-id="0887f-117">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="0887f-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="0887f-118">Le générateur Yeoman vous donne des instructions dans votre ligne de commande sur ce qu’il faut faire du projet, mais il les ignore et continue de suivre nos instructions.</span><span class="sxs-lookup"><span data-stu-id="0887f-118">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="0887f-119">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="0887f-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="0887f-120">Créez le projet.</span><span class="sxs-lookup"><span data-stu-id="0887f-120">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="0887f-121">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="0887f-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="0887f-122">Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0887f-122">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="0887f-123">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="0887f-123">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="0887f-124">Vous pouvez essayer le complément de fonction personnalisée dans Excel sur le Web ou Windows.</span><span class="sxs-lookup"><span data-stu-id="0887f-124">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="0887f-125">Vous serez peut-être invité à ouvrir le volet Office du complément, bien que ce soit facultatif.</span><span class="sxs-lookup"><span data-stu-id="0887f-125">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="0887f-126">Vous pouvez toujours exécuter vos fonctions personnalisées sans ouvrir le volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="0887f-126">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="0887f-127">Excel sur Windows</span><span class="sxs-lookup"><span data-stu-id="0887f-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="0887f-128">Pour tester votre complément dans Excel sous Windows, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="0887f-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="0887f-129">Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="0887f-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="0887f-130">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="0887f-130">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="0887f-131">Pour tester votre complément dans Excel sur le Web, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="0887f-131">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="0887f-132">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="0887f-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="0887f-133">Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="0887f-133">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="0887f-134">Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="0887f-134">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="0887f-135">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.</span><span class="sxs-lookup"><span data-stu-id="0887f-135">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="0887f-137">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="0887f-137">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="0887f-138">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="0887f-138">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="0887f-139">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="0887f-139">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="0887f-140">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="0887f-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="0887f-141">Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Yeoman contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="0887f-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="0887f-142">Le fichier **./manifest.xml** dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à `CONTOSO` l’espace de noms.</span><span class="sxs-lookup"><span data-stu-id="0887f-142">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="0887f-143">Dans votre classeur Excel, essayez la `ADD` fonction personnalisée en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="0887f-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="0887f-144">Sélectionnez une cellule et tapez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0887f-144">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="0887f-145">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0887f-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="0887f-146">Exécutez la `CONTOSO.ADD` fonction, en utilisant `10` des `200` nombres et comme paramètres d’entrée, en `=CONTOSO.ADD(10,200)` tapant la valeur dans la cellule et en appuyant sur entrée.</span><span class="sxs-lookup"><span data-stu-id="0887f-146">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="0887f-147">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="0887f-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="0887f-148">La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="0887f-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0887f-149">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="0887f-149">Next steps</span></span>

<span data-ttu-id="0887f-150">Félicitations, vous avez créé une fonction personnalisée dans un complément Excel !</span><span class="sxs-lookup"><span data-stu-id="0887f-150">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="0887f-151">Ensuite, créez un complément plus complexe avec la fonctionnalité de diffusion de données en continu.</span><span class="sxs-lookup"><span data-stu-id="0887f-151">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="0887f-152">Le lien suivant vous guide tout au long des étapes suivantes du didacticiel de complément Excel avec fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="0887f-152">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0887f-153">Didacticiel de complément de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="0887f-153">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="0887f-154">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0887f-154">See also</span></span>

* [<span data-ttu-id="0887f-155">Vue d’ensemble des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="0887f-155">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="0887f-156">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="0887f-156">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="0887f-157">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="0887f-157">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)