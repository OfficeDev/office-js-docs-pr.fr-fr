---
title: Déboguer des compléments Office sur un Mac
description: ''
ms.date: 07/29/2019
localization_priority: Priority
ms.openlocfilehash: 10b1181cab23252137df299736341c990978aa1d
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940681"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="b9442-102">Déboguer des compléments Office sur un Mac</span><span class="sxs-lookup"><span data-stu-id="b9442-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="b9442-p101">Étant donné que les compléments sont développés avec du code HTML et JavaScript, ils sont conçus pour fonctionner sur toutes les plateformes, mais il peut y avoir de subtiles différences dans le rendu du code HTML par les différents navigateurs. Cet article décrit la procédure de débogage des compléments qui s’exécutent sur un Mac.</span><span class="sxs-lookup"><span data-stu-id="b9442-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on a Mac. Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="b9442-105">Débogage avec l’inspecteur web Safari sur Mac</span><span class="sxs-lookup"><span data-stu-id="b9442-105">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="b9442-106">Si votre complément affiche une interface utilisateur dans un volet des tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l’aide de avec l’inspecteur web Safari.</span><span class="sxs-lookup"><span data-stu-id="b9442-106">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="b9442-107">Pour pouvoir déboguer des compléments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office version 16.9.1 (build 18012504) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9442-107">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="b9442-108">Si vous n’avez pas de build Mac Office, vous pouvez en obtenir une en rejoignant le [programme pour les développeurs Office 365](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="b9442-108">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="b9442-109">Pour commencer, ouvrez un terminal, puis définissez la propriété `OfficeWebAddinDeveloperExtras` pour l’application Office pertinente comme suit :</span><span class="sxs-lookup"><span data-stu-id="b9442-109">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="b9442-110">Ensuite, ouvrez l’application Office et[insérez votre complément](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="b9442-110">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="b9442-111">Cliquez sur le complément. Vous devriez voir l’option **Inspecter l’élément** s’afficher dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="b9442-111">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="b9442-112">Sélectionnez cette option pour afficher l’inspecteur dans lequel vous pouvez définir des points d’arrêt et déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="b9442-112">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b9442-113">Si vous essayez d’utiliser l’inspecteur et si la boîte de dialogue scintille, mettez Office à jour vers la dernière version.</span><span class="sxs-lookup"><span data-stu-id="b9442-113">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="b9442-114">Si cela ne résout pas le problème de scintillement, essayez la solution de contournement suivante :</span><span class="sxs-lookup"><span data-stu-id="b9442-114">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="b9442-115">Pour réduire la taille de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="b9442-115">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="b9442-116">Sélectionnez l’option **Inspecter l’élément** qui ouvre une nouvelle fenêtre.</span><span class="sxs-lookup"><span data-stu-id="b9442-116">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="b9442-117">Redimensionner la boîte de dialogue à sa taille d’origine.</span><span class="sxs-lookup"><span data-stu-id="b9442-117">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="b9442-118">Utiliser l’inspecteur comme requis.</span><span class="sxs-lookup"><span data-stu-id="b9442-118">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="b9442-119">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="b9442-119">Clearing the Office application's cache on a Mac or iPad</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
