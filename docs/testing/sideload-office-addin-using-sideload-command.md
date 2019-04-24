---
title: Chargement de versions test de compléments Office à l’aide de la commande sideload
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: dfa231374133ad857554afaf343362f1415788f4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449966"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="5c8be-102">Chargement de versions test de compléments Office pour les tester à l’aide de la **commande sideload**</span><span class="sxs-lookup"><span data-stu-id="5c8be-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="5c8be-103">La méthode « npm run sideload » fonctionne uniquement pour les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows ; et uniquement pour les projets de complément qui ont été créés dans l’outil [**yo office** ](https://github.com/OfficeDev/generator-office)et qui ont un script `sideload` dans la section `scripts` du fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="5c8be-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="5c8be-104">(Les projets qui ont été créés dans les versions antérieures de **yo office** n’ont pas ce script non plus.) Si votre projet a été créé avec Visual Studio ou n’a pas le script sideload , vous pouvez charger une version test sur Windows en suivant la méthode décrite dans l’article relatif au [chargement de version test d’un complément Office à partir d’un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="5c8be-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="5c8be-105">Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :</span><span class="sxs-lookup"><span data-stu-id="5c8be-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="5c8be-106">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="5c8be-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="5c8be-107">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="5c8be-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="5c8be-108">Chargement de version test des compléments Outlook pour les tester</span><span class="sxs-lookup"><span data-stu-id="5c8be-108">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="5c8be-109">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="5c8be-109">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="5c8be-110">Modifiez les répertoires vers la racine du dossier de votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="5c8be-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="5c8be-111">Exécutez la commande suivante pour démarrer une instance du serveur web local sur le port 3000 et mettre en service votre projet de complément : « **npm exécuter début** »</span><span class="sxs-lookup"><span data-stu-id="5c8be-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="5c8be-112">Ouvrez une deuxième invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="5c8be-112">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="5c8be-113">Modifiez les répertoires vers la racine du dossier de votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="5c8be-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="5c8be-114">Exécutez la commande suivante pour démarrer l’application hôte (par exemple, Excel, Word) et inscrire votre complément dans l’application hôte : « **npm run sideloadr** »</span><span class="sxs-lookup"><span data-stu-id="5c8be-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="5c8be-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5c8be-115">See also</span></span>

- [<span data-ttu-id="5c8be-116">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="5c8be-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="5c8be-117">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="5c8be-117">Publish your Office Add-in</span></span>](../publish/publish.md)
