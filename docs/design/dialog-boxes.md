---
title: Boîtes de dialogue dans les compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f18f603d76a902bdce56152ecb3f63bbafad56fb
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945749"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="483f2-102">Boîtes de dialogue dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="483f2-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="483f2-p101">Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Office. Vous pouvez utiliser les boîtes de dialogue afin de fournir un espace supplémentaire sur l’écran pour les tâches comme les pages de connexion impossibles à ouvrir directement dans un volet des tâches, ou pour les demandes de confirmation d’une action effectuée par un utilisateur, ou pour afficher des vidéos qui peuvent être trop petites si confinées à un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="483f2-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="483f2-105">*Figure 1. Mise en page type pour une boîte de dialogue*</span><span class="sxs-lookup"><span data-stu-id="483f2-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Exemple d’image affichant une mise en page par défaut pour une boîte de dialogue](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="483f2-107">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="483f2-107">Best practices</span></span>

|<span data-ttu-id="483f2-108">**À faire**</span><span class="sxs-lookup"><span data-stu-id="483f2-108">**Do**</span></span>|<span data-ttu-id="483f2-109">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="483f2-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="483f2-110">Inclure un titre descriptif qui inclut le nom de votre complément, ainsi que la tâche en cours.</span><span class="sxs-lookup"><span data-stu-id="483f2-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="483f2-111">Ne pas ajouter le nom de votre société au titre.</span><span class="sxs-lookup"><span data-stu-id="483f2-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="483f2-112">Ne pas ouvrir une boîte de dialogue, sauf si le scénario l’exige.</span><span class="sxs-lookup"><span data-stu-id="483f2-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="483f2-113">Implémentation</span><span class="sxs-lookup"><span data-stu-id="483f2-113">Implementation</span></span>

<span data-ttu-id="483f2-114">Pour voir un exemple relatif à l’implémentation d’une boîte de dialogue, consultez [Exemple d’API de boîte de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="483f2-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="483f2-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="483f2-115">See also</span></span>

- [<span data-ttu-id="483f2-116">Exemple de modèle UX</span><span class="sxs-lookup"><span data-stu-id="483f2-116">UX Pattern Sample</span></span>](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [<span data-ttu-id="483f2-117">Ressources de développement GitHub</span><span class="sxs-lookup"><span data-stu-id="483f2-117">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="483f2-118">Objet Dialogue</span><span class="sxs-lookup"><span data-stu-id="483f2-118">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js)


