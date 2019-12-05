---
title: Développement d’un complément Office avec Visual Studio Code
description: Comment développer un complément Office avec Visual Studio Code
ms.date: 12/02/2019
localization_priority: Priority
ms.openlocfilehash: a18d8a74ff269b32e83c836b06629850873e507b
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670485"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="1cd8a-103">Développement d’un complément Office avec Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="1cd8a-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="1cd8a-104">Cet article explique comment utiliser [Visual Studio Code (VS Code)](https://code.visualstudio.com) pour développer votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1cd8a-105">Pour plus d’informations sur l’utilisation de Visual Studio pour créer un complément Office, voir [Créer et déboguer des compléments Office dans Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="1cd8a-105">For information about using Visual Studio to create an Office Add-in, see [Create and debug Office Add-ins in Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1cd8a-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="1cd8a-106">Prerequisites</span></span>

- [<span data-ttu-id="1cd8a-107">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="1cd8a-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="1cd8a-108">Créez le projet de complément à l’aide du générateur Yeoman</span><span class="sxs-lookup"><span data-stu-id="1cd8a-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="1cd8a-109">Si vous utilisez le VS Code comme environnement de développement intégré (IDE), vous devez créer le projet de complément Office avec le [genérateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office). Le générateur Yeoman crée un projet Node js qui peut être géré avec VS Code ou n’importe quel autre éditeur.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="1cd8a-110">Pour créer un complément Office avec le générateur Yeoman, suivez les instructions dans le [démarrage rapide de 5 minutes](../index.md) qui correspond au type de complément que vous voulez créer.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="1cd8a-111">Développer le complément à l’aide de VS Code</span><span class="sxs-lookup"><span data-stu-id="1cd8a-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="1cd8a-112">Lorsque le générateur Yeoman a terminé de créer le projet de complément, ouvrez le dossier racine du projet avec VS Code.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="1cd8a-113">Dans Windows, vous pouvez accéder au répertoire racine du projet via la ligne de commande, puis entrer `code .` pour ouvrir ce dossier dans VS Code.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="1cd8a-114">Sur Mac, vous devez [ajouter la commande `code` au chemin d’accès](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) avant de pouvoir utiliser cette commande pour ouvrir le dossier de projet dans VS Code.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="1cd8a-115">Le générateur Yeoman crée un complément de base avec une fonctionnalité limitée.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="1cd8a-116">Vous pouvez personnaliser le complément en modifiant le [manifeste](add-in-manifests.md), HTML, JavaScript ou TypeScript et des fichiers CSS dans VS Code.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="1cd8a-117">Pour obtenir une description générale de la structure de projet et des fichiers dans le projet de complément que le générateur Yeoman crée, consultez les instructions du générateur Yeoman dans le [démarrage rapide de 5 minutes](../index.md) qui correspond au type de complément que vous avez créé.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="1cd8a-118">Tester et déboguer le complément</span><span class="sxs-lookup"><span data-stu-id="1cd8a-118">To run and debug the add-in</span></span>

<span data-ttu-id="1cd8a-119">Les méthodes de test, de débogage et de résolution des problèmes liés aux compléments Office varient selon la plateforme.</span><span class="sxs-lookup"><span data-stu-id="1cd8a-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="1cd8a-120">Pour plus d’informations, voir [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="1cd8a-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="1cd8a-121">Publier le complément</span><span class="sxs-lookup"><span data-stu-id="1cd8a-121">Publish the add-in.</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="1cd8a-122">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1cd8a-122">See also</span></span>

- [<span data-ttu-id="1cd8a-123">Démarrages rapides de 5?minutes</span><span class="sxs-lookup"><span data-stu-id="1cd8a-123">5-Minute Quick Starts</span></span>](../index.md)
- <span data-ttu-id="1cd8a-124">[Explorer l’API JavaScript Office à l’aide de Script Lab](../overview/explore-with-script-lab.md)</span><span class="sxs-lookup"><span data-stu-id="1cd8a-124">To learn more, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).</span></span>
- [<span data-ttu-id="1cd8a-125">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="1cd8a-125">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="1cd8a-126">Déploiement et publication de votre complément Office</span><span class="sxs-lookup"><span data-stu-id="1cd8a-126">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)