---
title: Chargement de versions test de compléments Office à l’aide de la commande sideload
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: 69d39c2736312653b5a362aefccd41629e6e3555
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619076"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="c10b8-102">Chargement indépendant de compléments Office pour les tester à l’aide de la commande sideload</span><span class="sxs-lookup"><span data-stu-id="c10b8-102">Sideload Office Add-ins for testing using the sideload command</span></span>
 
> [!NOTE]
> <span data-ttu-id="c10b8-103">La technique de chargement indépendant décrite dans cet article est uniquement valide pour :</span><span class="sxs-lookup"><span data-stu-id="c10b8-103">The sideloading technique described in this article is only valid for:</span></span>
> 
> - <span data-ttu-id="c10b8-104">Les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows.</span><span class="sxs-lookup"><span data-stu-id="c10b8-104">Excel, Word, and PowerPoint add-ins that run on Windows</span></span>
> 
> - <span data-ttu-id="c10b8-105">Les projets de complément créés avec le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office) et disposant d’un script `sideload` dans la section `scripts` du fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="c10b8-105">Add-in projects that were created with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="c10b8-106">(Ce script n’est pas présent dans les projets créés avec d’anciennes versions du générateur Yeoman pour compléments Office).</span><span class="sxs-lookup"><span data-stu-id="c10b8-106">(Projects that were created with older versions of the Yeoman generator for Office Add-ins will not have this script.)</span></span>
 
<span data-ttu-id="c10b8-107">Pour charger indépendamment votre complément à l’aide du script `sideload` fourni par le générateur Yeoman pour compléments Office, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="c10b8-107">To sideload your add-in by using the `sideload` script that the Yeoman generator for Office Add-ins provides, complete the following steps:</span></span>

1. <span data-ttu-id="c10b8-108">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="c10b8-108">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="c10b8-109">Modifiez les répertoires vers la racine du dossier de votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="c10b8-109">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="c10b8-110">Exécutez la commande suivante pour démarrer une instance du serveur web local sur le port 3000 et mettre en service votre projet de complément : `npm run start`</span><span class="sxs-lookup"><span data-stu-id="c10b8-110">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "`npm run start`"</span></span>

4. <span data-ttu-id="c10b8-111">Ouvrez une deuxième invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="c10b8-111">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="c10b8-112">Modifiez les répertoires vers la racine du dossier de votre projet de complément.</span><span class="sxs-lookup"><span data-stu-id="c10b8-112">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="c10b8-113">Exécutez la commande suivante pour démarrer l’application hôte (par exemple, Excel, Word) et inscrire votre complément dans l’application hôte : `npm run sideload`</span><span class="sxs-lookup"><span data-stu-id="c10b8-113">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "`npm run sideload`"</span></span>

<span data-ttu-id="c10b8-114">Si votre projet de complément a été créé avec Visual Studio ou n’a pas le script sideload , vous pouvez le charger indépendamment sur Windows en suivant la méthode décrite dans l’article relatif au [chargement indépendant d’un complément Office à partir d’un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="c10b8-114">(Projects that were created with older versions of yo office also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="c10b8-115">Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour plus d’informations sur le chargement indépendant de votre complément :</span><span class="sxs-lookup"><span data-stu-id="c10b8-115">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
 
- [<span data-ttu-id="c10b8-116">Chargement de version test des compléments Office dans Office Online pour les tester</span><span class="sxs-lookup"><span data-stu-id="c10b8-116">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="c10b8-117">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="c10b8-117">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="c10b8-118">Chargement de version test des compléments Outlook pour les tester</span><span class="sxs-lookup"><span data-stu-id="c10b8-118">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a><span data-ttu-id="c10b8-119">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c10b8-119">See also</span></span>

- [<span data-ttu-id="c10b8-120">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="c10b8-120">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="c10b8-121">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="c10b8-121">Publish your Office Add-in</span></span>](../publish/publish.md)
