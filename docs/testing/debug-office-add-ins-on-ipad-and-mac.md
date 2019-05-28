---
title: Déboguer des compléments Office sur un Mac
description: ''
ms.date: 05/21/2019
localization_priority: Priority
ms.openlocfilehash: 0505dcc49ea98040f1c4891621c8e30a8cbeaff4
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432277"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="ee3b9-102">Déboguer des compléments Office sur un Mac</span><span class="sxs-lookup"><span data-stu-id="ee3b9-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="ee3b9-p101">Vous pouvez utiliser Visual Studio pour développer et déboguer des compléments sous Windows, mais vous ne pouvez pas l’utiliser pour déboguer des compléments sur un Mac. Étant donné que les compléments sont développés avec du code HTML et JavaScript, ils sont conçus pour fonctionner sur toutes les plateformes, mais il peut y avoir de subtiles différences dans le rendu du code HTML par les différents navigateurs. Cet article décrit la procédure de débogage des compléments qui s’exécutent sur un Mac.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on a Mac. Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="ee3b9-106">Débogage avec l’inspecteur web Safari sur Mac</span><span class="sxs-lookup"><span data-stu-id="ee3b9-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="ee3b9-107">Si votre complément affiche une interface utilisateur dans un volet des tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l’aide de avec l’inspecteur web Safari.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="ee3b9-108">Pour pouvoir déboguer des compléments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office version 16.9.1 (build 18012504) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="ee3b9-109">Si vous n’avez pas de build Mac Office, vous pouvez en obtenir une en rejoignant le [programme pour les développeurs Office 365](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="ee3b9-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="ee3b9-110">Pour commencer, ouvrez un terminal, puis définissez la propriété `OfficeWebAddinDeveloperExtras` pour l’application Office pertinente comme suit :</span><span class="sxs-lookup"><span data-stu-id="ee3b9-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="ee3b9-111">Ensuite, ouvrez l’application Office et[insérez votre complément](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="ee3b9-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="ee3b9-112">Cliquez sur le complément. Vous devriez voir l’option **Inspecter l’élément** s’afficher dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="ee3b9-113">Sélectionnez cette option pour afficher l’inspecteur dans lequel vous pouvez définir des points d’arrêt et déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="ee3b9-114">Si vous essayez d’utiliser l’inspecteur et si la boîte de dialogue scintille, mettez Office à jour vers la dernière version.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="ee3b9-115">Si cela ne résout pas le problème de scintillement, essayez la solution de contournement suivante :</span><span class="sxs-lookup"><span data-stu-id="ee3b9-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="ee3b9-116">Pour réduire la taille de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="ee3b9-117">Sélectionnez l’option **Inspecter l’élément** qui ouvre une nouvelle fenêtre.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="ee3b9-118">Redimensionner la boîte de dialogue à sa taille d’origine.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="ee3b9-119">Utiliser l’inspecteur comme requis.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="ee3b9-120">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="ee3b9-120">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="ee3b9-p105">Les compléments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances. En règle générale, vous pouvez effacer le cache en rechargeant le complément. En présence de plusieurs compléments dans le même document, il se peut que le processus d’effacement automatique du cache lors du rechargement ne fonctionne pas systématiquement.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-p105">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="ee3b9-124">Sur un ordinateur Mac, vous pouvez effacer le cache manuellement en supprimant le contenu du dossier `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="ee3b9-124">On a Mac, you can clear the cache manually by deleting everything in the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
