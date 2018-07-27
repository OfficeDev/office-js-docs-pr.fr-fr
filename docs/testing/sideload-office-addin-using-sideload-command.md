---
title: Charger une version test des compléments Office à l'aide de la commande de chargement indépendant
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279359"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="0dc5f-102">Chargez une version test des compléments Office à l'aide de la **commande de chargement indépendant**</span><span class="sxs-lookup"><span data-stu-id="0dc5f-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="0dc5f-103">La méthode « npm run sideload » ne fonctionne que pour les compléments Excel, Word et PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="0dc5f-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

1. <span data-ttu-id="0dc5f-104">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="0dc5f-104">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="0dc5f-105">Modifiez les répertoires à la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="0dc5f-105">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="0dc5f-106">Exécutez la commande suivante pour démarrer une instance de serveur Web local sur le port 3000 afin de servir votre projet de complément :**« npm run start »**</span><span class="sxs-lookup"><span data-stu-id="0dc5f-106">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="0dc5f-107">Ouvrez une nouvelle invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="0dc5f-107">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="0dc5f-108">Changez les répertoires à la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="0dc5f-108">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="0dc5f-109">Exécutez la commande suivante pour démarrer l'application hôte (par exemple Excel, Word) et enregistrez votre complément dans l'application hôte :**« npm run sideload »**</span><span class="sxs-lookup"><span data-stu-id="0dc5f-109">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="0dc5f-110">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0dc5f-110">See also</span></span>

- [<span data-ttu-id="0dc5f-111">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="0dc5f-111">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="0dc5f-112">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="0dc5f-112">Publish your Office Add-in</span></span>](../publish/publish.md)