---
ms.date: 05/17/2020
description: Découvrez l'exécution de fonctions personnalisées, les boutons du ruban et le code du volet des tâches dans un runtime JavaScript identique pour coordonner des scénarios dans votre complément.
title: Exécuter le code de votre complément dans un Runtime JavaScript partagé
localization_priority: Priority
ms.openlocfilehash: afb07c5223e26ba1e1adbf40c7a4b2e4f7c06349
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275930"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a><span data-ttu-id="b45db-103">Vue d’ensemble : exécuter le code de votre complément dans un Runtime JavaScript partagé</span><span class="sxs-lookup"><span data-stu-id="b45db-103">Overview: Run your add-in code in a shared JavaScript runtimes</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="b45db-104">Lors de l’exécution d’Excel sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts.</span><span class="sxs-lookup"><span data-stu-id="b45db-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="b45db-105">Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b45db-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="b45db-106">Vous pouvez toutefois configurer votre complément Excel pour partager un code dans le même runtime JavaScript (également appelé runtime partagé).</span><span class="sxs-lookup"><span data-stu-id="b45db-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="b45db-107">Vous pouvez ainsi améliorer la coordination dans votre complément et accéder au volet des tâches DOM et CORS à partir de toutes les parties de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b45db-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="b45db-108">La configuration d’un runtime partagé permet les scénarios suivants :</span><span class="sxs-lookup"><span data-stu-id="b45db-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="b45db-109">Votre complément dispose d'un DOM partagé auquel le ruban, le volet des tâches et les fonctions personnalisées peuvent accéder.</span><span class="sxs-lookup"><span data-stu-id="b45db-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="b45db-110">Vos fonctions personnalisées bénéficieront d'une prise en charge complète de CORS.</span><span class="sxs-lookup"><span data-stu-id="b45db-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="b45db-111">Vos fonctions personnalisées peuvent appeler les API Office.js pour lire les données d’un document feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b45db-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="b45db-112">Votre complément peut exécuter un code dès que le document est ouvert.</span><span class="sxs-lookup"><span data-stu-id="b45db-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="b45db-113">Votre complément peut continuer à exécuter un code lorsque le volet des tâches est fermé.</span><span class="sxs-lookup"><span data-stu-id="b45db-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="b45db-114">Lorsque vous exécutez des fonctions personnalisées dans un runtime partagé avec le volet des tâches, celui-ci s’exécute dans une instance de navigateur sur différentes plateformes, tel qu'expliqué dans [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md). En outre, les boutons affichés sur le ruban par votre complément Excel s’exécutent dans le même runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="b45db-114">When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="b45db-115">L’image ci-après présente l'exécution des fonctions personnalisées, de interface utilisateur du ruban et du code du volet des tâches dans le même runtime JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b45db-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Fonctions personnalisées en cours d’exécution dans un runtime partagé avec des boutons du ruban et le volet Office dans Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a><span data-ttu-id="b45db-117">Configurer un runtime partagé</span><span class="sxs-lookup"><span data-stu-id="b45db-117">Set up a shared runtime</span></span>

<span data-ttu-id="b45db-118">Consultez la rubrique [Configuring a Shared Runtime article](./configure-your-add-in-to-use-a-shared-runtime.md) pour apprendre à configurer vos fonctions personnalisées afin d’utiliser un runtime partagé.</span><span class="sxs-lookup"><span data-stu-id="b45db-118">See the [configuring a shared runtime article](./configure-your-add-in-to-use-a-shared-runtime.md) to learn how to set up your custom functions to use a shared runtime.</span></span>

### <a name="debugging"></a><span data-ttu-id="b45db-119">Débogage</span><span class="sxs-lookup"><span data-stu-id="b45db-119">Debugging</span></span>

<span data-ttu-id="b45db-120">Lors de l’utilisation d’un runtime partagé, vous ne pouvez pas utiliser Visual Studio Code pour déboguer des fonctions personnalisées dans Excel sur Windows à cette date.</span><span class="sxs-lookup"><span data-stu-id="b45db-120">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="b45db-121">Vous devez plutôt utiliser des outils de développement.</span><span class="sxs-lookup"><span data-stu-id="b45db-121">You'll need to use developer tools instead.</span></span> <span data-ttu-id="b45db-122">Pour plus d'informations, voir le [Débogage des compléments avec les outils de développement sur Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span><span class="sxs-lookup"><span data-stu-id="b45db-122">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="b45db-123">Faites nous part de vos commentaires</span><span class="sxs-lookup"><span data-stu-id="b45db-123">Give us feedback</span></span>

<span data-ttu-id="b45db-124">Nous aimerions connaître votre avis concernant cette fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="b45db-124">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="b45db-125">Si vous trouvez des bogues, des problèmes ou si vous avez des questions relatives à cette fonctionnalité, faites-le nous savoir en créant un problème GitHub dans le [référentiel Office-js](https://github.com/OfficeDev/office-js).</span><span class="sxs-lookup"><span data-stu-id="b45db-125">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="b45db-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b45db-126">See also</span></span>

- [<span data-ttu-id="b45db-127">Didacticiel : partager des données et des événements entre des fonctions personnalisées Excel et le volet Office</span><span class="sxs-lookup"><span data-stu-id="b45db-127">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="b45db-128">Appeler des API Excel à partir de votre fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="b45db-128">Call Excel APIs from your custom function</span></span>](call-excel-apis-from-custom-function.md)
