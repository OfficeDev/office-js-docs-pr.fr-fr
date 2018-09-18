---
title: Charger une version test des compléments Office à l'aide de la commande de chargement indépendant
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 1ab0277493f2899adb479c2f24b1635a881af3cc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944040"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="476c6-102">Chargez une version test des compléments Office à l'aide de la **commande de chargement indépendant**</span><span class="sxs-lookup"><span data-stu-id="476c6-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="476c6-p101">La méthode « npm exécuter sideload » ne fonctionne que pour les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows ; et uniquement pour les projets de compléments qui ont été créés avec l'outil [**yo office**](https://github.com/OfficeDev/generator-office) et qui ont un script `sideload` dans la section `scripts` du fichier package.json. (Les projets qui ont été créées avec les versions antérieures de **yo office** n’ont pas non plus ce script.) Si votre projet a été créé avec Visual Studio ou n’a pas le script sideload, vous pouvez le charger en version test sur Windows avec la méthode décrite dans [Chargement de la version test d'un complément Office depuis un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="476c6-p101">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file. (Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="476c6-105">Si ce n'est pas un complément Word, Excel ou PowerPoint sous Windows que vous testez, consultez une des rubriques suivantes pour charger la version test de votre complément :</span><span class="sxs-lookup"><span data-stu-id="476c6-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="476c6-106">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="476c6-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="476c6-107">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="476c6-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="476c6-108">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="476c6-108">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="476c6-109">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="476c6-109">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="476c6-110">Modifiez les répertoires à la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="476c6-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="476c6-111">Exécutez la commande suivante pour démarrer une instance de serveur Web local sur le port 3000 afin de servir votre projet de complément :**« npm run start »**</span><span class="sxs-lookup"><span data-stu-id="476c6-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="476c6-112">Ouvrez une nouvelle invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="476c6-112">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="476c6-113">Changez les répertoires à la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="476c6-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="476c6-114">Exécutez la commande suivante pour démarrer l'application hôte (par exemple Excel, Word) et enregistrez votre complément dans l'application hôte :**« npm run sideload »**</span><span class="sxs-lookup"><span data-stu-id="476c6-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="476c6-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="476c6-115">See also</span></span>

- [<span data-ttu-id="476c6-116">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="476c6-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="476c6-117">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="476c6-117">Publish your Office Add-in</span></span>](../publish/publish.md)