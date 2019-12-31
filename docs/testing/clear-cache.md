---
title: Vider le cache Office
description: Découvrez comment effacer le cache Office sur votre ordinateur.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915056"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="aeb40-103">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="aeb40-103">Clear the Office cache</span></span>

<span data-ttu-id="aeb40-104">Vous pouvez supprimer un complément que vous avez précédemment chargé sur Windows, Mac ou iOS en vidant le cache Office sur votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="aeb40-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="aeb40-105">En outre, si vous apportez des modifications au manifeste de votre complément (par exemple, vous mettez à jour le nom des fichiers d’icônes ou de texte de commandes du complément), videz le cache Office, puis rechargez le complément à l’aide d’un manifeste mis à jour.</span><span class="sxs-lookup"><span data-stu-id="aeb40-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="aeb40-106">Cette action permettra à Office d’afficher le complément tel que décrit par le manifeste mis à jour.</span><span class="sxs-lookup"><span data-stu-id="aeb40-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="aeb40-107">Vider le cache Office sur Windows</span><span class="sxs-lookup"><span data-stu-id="aeb40-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="aeb40-108">Pour vider le cache Office sur Windows, supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="aeb40-108">To clear the Office cache on Windows, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="aeb40-109">Vider le cache Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="aeb40-109">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="aeb40-110">Vider le cache Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="aeb40-110">Clear the Office cache on iOS</span></span>

<span data-ttu-id="aeb40-111">Pour vider le cache Office sur iOS, appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement.</span><span class="sxs-lookup"><span data-stu-id="aeb40-111">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="aeb40-112">Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="aeb40-112">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="aeb40-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="aeb40-113">See also</span></span>

- [<span data-ttu-id="aeb40-114">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="aeb40-114">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="aeb40-115">Valider le manifeste d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="aeb40-115">Validate an Office Add-in manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="aeb40-116">Déboguer votre complément avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="aeb40-116">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="aeb40-117">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="aeb40-117">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="aeb40-118">Débogage des compléments Office</span><span class="sxs-lookup"><span data-stu-id="aeb40-118">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)