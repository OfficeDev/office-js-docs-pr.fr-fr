---
ms.date: 05/02/2019
description: Développement de fonctions personnalisées dans le Guide de démarrage rapide d’Excel.
title: Démarrage rapide des fonctions personnalisées
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8eb2630526ce939273024eebd533bd99fa5e94a1
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619893"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="30eaf-103">Prise en main du développement de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="30eaf-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="30eaf-104">Avec les fonctions personnalisées, les développeurs peuvent désormais ajouter de nouvelles fonctions à Excel en les définissant en JavaScript ou en une machine à écrire dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="30eaf-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="30eaf-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe `SUM()`quelle fonction native dans Excel, comme.</span><span class="sxs-lookup"><span data-stu-id="30eaf-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="30eaf-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="30eaf-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="30eaf-107">Excel pour Windows (version 64 bits 1810 ou ultérieure) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="30eaf-107">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="30eaf-108">Rejoignez le[programme Office Insider](https://products.office.com/office-insider)(\*\* niveau\*\*Insider, anciennement appelé « Insider Fast »)</span><span class="sxs-lookup"><span data-stu-id="30eaf-108">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="30eaf-109">Création de votre premier projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="30eaf-109">Build your first custom functions project</span></span>

<span data-ttu-id="30eaf-110">Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="30eaf-110">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="30eaf-111">Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="30eaf-111">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="30eaf-112">Dans un dossier de votre choix, exécutez la commande suivante, puis répondez aux invites comme suit.</span><span class="sxs-lookup"><span data-stu-id="30eaf-112">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="30eaf-113">**Sélectionnez un type de projet :** `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="30eaf-113">**Choose a project type:** `Excel Custom Functions Add-in project (...)`</span></span>
    - <span data-ttu-id="30eaf-114">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="30eaf-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="30eaf-115">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="30eaf-115">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/yo-office-excel-cf.png)

    <span data-ttu-id="30eaf-117">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="30eaf-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="30eaf-118">Le générateur Yeoman vous donne des instructions dans votre ligne de commande sur ce qu’il faut faire du projet, mais il les ignore et continue de suivre nos instructions.</span><span class="sxs-lookup"><span data-stu-id="30eaf-118">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="30eaf-119">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="30eaf-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="30eaf-120">Créez le projet.</span><span class="sxs-lookup"><span data-stu-id="30eaf-120">Build the project.</span></span> <span data-ttu-id="30eaf-121">Cette opération installe également les certificats dont votre projet a besoin pour fonctionner correctement.</span><span class="sxs-lookup"><span data-stu-id="30eaf-121">This will also install certificates that your project needs in order to function properly.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="30eaf-122">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="30eaf-122">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="30eaf-123">Vous pouvez tester le complément de fonction personnalisée dans Excel pour Windows ou Excel online.</span><span class="sxs-lookup"><span data-stu-id="30eaf-123">You can try out the custom function add-in in Excel for Windows or Excel Online.</span></span> <span data-ttu-id="30eaf-124">Vous serez peut-être invité à ouvrir le volet Office du complément, bien que ce soit facultatif.</span><span class="sxs-lookup"><span data-stu-id="30eaf-124">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="30eaf-125">Vous pouvez toujours exécuter vos fonctions personnalisées sans ouvrir le volet Office de votre complément.</span><span class="sxs-lookup"><span data-stu-id="30eaf-125">You can still run your custom functions without opening your add-in's task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="30eaf-126">Les compléments Office doivent utiliser le protocole HTTPs, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="30eaf-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="30eaf-127">Si vous êtes invité à installer un certificat après l’avoir exécuté `npm run start:desktop`, acceptez l’invite pour installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="30eaf-127">If you are prompted to install a certificate after you run `npm run start:desktop`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="30eaf-128">Excel pour Windows</span><span class="sxs-lookup"><span data-stu-id="30eaf-128">Excel for Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="30eaf-129">Pour tester votre complément dans Excel pour Windows, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="30eaf-129">To test your add-in in Excel for Windows, run the following command.</span></span> <span data-ttu-id="30eaf-130">Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="30eaf-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="30eaf-131">Excel Online</span><span class="sxs-lookup"><span data-stu-id="30eaf-131">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="30eaf-132">Pour tester votre complément dans Excel Online, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="30eaf-132">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="30eaf-133">Lorsque vous exécutez cette commande, le serveur Web local démarre.</span><span class="sxs-lookup"><span data-stu-id="30eaf-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> <span data-ttu-id="30eaf-134">Les compléments Office doivent utiliser le protocole HTTPs, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="30eaf-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="30eaf-135">Si vous êtes invité à installer un certificat après l’avoir exécuté `npm run start:web`, acceptez l’invite pour installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="30eaf-135">If you are prompted to install a certificate after you run `npm run start:web`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

<span data-ttu-id="30eaf-136">Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel online.</span><span class="sxs-lookup"><span data-stu-id="30eaf-136">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="30eaf-137">Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="30eaf-137">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="30eaf-138">Dans Excel Online, sélectionnez l’onglet **Insérer**, puis **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="30eaf-138">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Insérer un ruban dans Excel Online avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="30eaf-140">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="30eaf-140">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="30eaf-141">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="30eaf-141">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="30eaf-142">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="30eaf-142">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="30eaf-143">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="30eaf-143">Try out a prebuilt custom function</span></span>

<span data-ttu-id="30eaf-144">Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Yeoman contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="30eaf-144">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="30eaf-145">Le fichier **./manifest.xml** dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à `CONTOSO` l’espace de noms.</span><span class="sxs-lookup"><span data-stu-id="30eaf-145">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="30eaf-146">Dans votre classeur Excel, essayez la `ADD` fonction personnalisée en procédant comme suit:</span><span class="sxs-lookup"><span data-stu-id="30eaf-146">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="30eaf-147">Sélectionnez une cellule et tapez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="30eaf-147">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="30eaf-148">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="30eaf-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="30eaf-149">Exécutez la `CONTOSO.ADD` fonction, en utilisant `10` des `200` nombres et comme paramètres d’entrée, en `=CONTOSO.ADD(10,200)` tapant la valeur dans la cellule et en appuyant sur entrée.</span><span class="sxs-lookup"><span data-stu-id="30eaf-149">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="30eaf-150">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="30eaf-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="30eaf-151">La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="30eaf-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="30eaf-152">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="30eaf-152">Next steps</span></span>

<span data-ttu-id="30eaf-153">Félicitations, vous avez créé une fonction personnalisée dans un complément Excel!</span><span class="sxs-lookup"><span data-stu-id="30eaf-153">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="30eaf-154">Ensuite, créez un complément plus complexe avec la fonctionnalité de diffusion de données en continu.</span><span class="sxs-lookup"><span data-stu-id="30eaf-154">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="30eaf-155">Le lien suivant vous guide tout au long des étapes suivantes du didacticiel de complément Excel avec fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="30eaf-155">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="30eaf-156">Didacticiel de complément de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="30eaf-156">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="30eaf-157">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="30eaf-157">See also</span></span>

* [<span data-ttu-id="30eaf-158">Vue d’ensemble des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="30eaf-158">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="30eaf-159">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="30eaf-159">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="30eaf-160">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="30eaf-160">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="30eaf-161">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="30eaf-161">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
