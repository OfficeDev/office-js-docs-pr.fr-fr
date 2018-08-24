---
title: Charger une version test des compléments Office à l'aide de la commande de chargement indépendant
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 3aacfdb09f362ea10ba0e2393caca335fe4c04c6
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925100"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="18959-102">Chargez une version test des compléments Office à l'aide de la **commande de chargement indépendant**</span><span class="sxs-lookup"><span data-stu-id="18959-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="18959-103">La méthode « npm run sideload » fonctionne uniquement pour les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows ; et uniquement pour les projets de complément créés avec l’outil [**Yo Office**](https://github.com/OfficeDev/generator-office) et disposant d’un `sideload` script dans la section `scripts` du fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="18959-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="18959-104">(Les projets créés avec des versions antérieures de **Yo Office** ne disposent pas de ce script.) Si votre projet a été créé avec Visual Studio ou ne dispose pas du script sideload, vous pouvez en charger la version test sur Windows en suivant la méthode décrite dans [Chargement d’une version test de complément Office à partir d’un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="18959-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="18959-105">Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :</span><span class="sxs-lookup"><span data-stu-id="18959-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="18959-106">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="18959-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="18959-107">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="18959-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="18959-108">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="18959-108">Sideload Outlook Add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="18959-109">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="18959-109">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="18959-110">Modifiez les répertoires à la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="18959-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="18959-111">Exécutez la commande suivante pour démarrer une instance de serveur Web local sur le port 3000 afin de servir votre projet de complément :**« npm run start »**</span><span class="sxs-lookup"><span data-stu-id="18959-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="18959-112">Ouvrez une nouvelle invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="18959-112">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="18959-113">Changez les répertoires à la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="18959-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="18959-114">Exécutez la commande suivante pour démarrer l'application hôte (par exemple Excel, Word) et enregistrez votre complément dans l'application hôte :**« npm run sideload »**</span><span class="sxs-lookup"><span data-stu-id="18959-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="18959-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="18959-115">See also</span></span>

- [<span data-ttu-id="18959-116">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="18959-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="18959-117">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="18959-117">Publish your Office Add-in</span></span>](../publish/publish.md)