---
title: Test et débogage de compléments Office
description: ''
ms.date: 11/24/2017
localization_priority: Priority
ms.openlocfilehash: 7ffa281807ca1541f8ebcc5f722c1043db115509
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388835"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="adbff-102">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="adbff-102">Test and debug Office Add-ins</span></span>

<span data-ttu-id="adbff-103">Cette section contient des recommandations sur les tests, le débogage et la résolution des problèmes avec les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="adbff-103">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="adbff-104">Chargement de version test d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="adbff-104">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="adbff-105">Vous pouvez utiliser le chargement de version test pour installer un complément Office sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="adbff-105">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog.</span></span> <span data-ttu-id="adbff-106">La procédure de chargement de version test d’un complément varie en fonction de la plateforme et, dans certains cas, du produit.</span><span class="sxs-lookup"><span data-stu-id="adbff-106">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="adbff-107">Les articles suivants décrivent comment charger une version test de compléments Office sur une plateforme spécifique ou dans un produit spécifique :</span><span class="sxs-lookup"><span data-stu-id="adbff-107">The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="adbff-108">Chargement de version test des compléments Office sur Windows</span><span class="sxs-lookup"><span data-stu-id="adbff-108">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="adbff-109">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="adbff-109">Sideload Office Add-ins in Office Online</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="adbff-110">Chargement de version test de compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="adbff-110">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="adbff-111">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="adbff-111">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="adbff-112">Débogage d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="adbff-112">Debug an Office Add-in</span></span>

<span data-ttu-id="adbff-113">La procédure pour le débogage d’un complément Office varie également selon la plateforme.</span><span class="sxs-lookup"><span data-stu-id="adbff-113">The procedure for debugging an Office Add-in varies by platform as well.</span></span> <span data-ttu-id="adbff-114">Chacun des articles suivants décrit comment déboguer des compléments Office sur une plateforme spécifique :</span><span class="sxs-lookup"><span data-stu-id="adbff-114">Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="adbff-115">Attacher un débogueur à partir du volet Office (sur Windows)</span><span class="sxs-lookup"><span data-stu-id="adbff-115">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="adbff-116">Débogage de compléments avec les outils de développement F12 sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="adbff-116">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="adbff-117">Débogage de compléments dans Office Online</span><span class="sxs-lookup"><span data-stu-id="adbff-117">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="adbff-118">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="adbff-118">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="adbff-119">Validation d’un manifeste de complément Office</span><span class="sxs-lookup"><span data-stu-id="adbff-119">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="adbff-120">Pour savoir comment valider le fichier manifeste qui décrit votre complément Office et résoudre des problèmes avec le fichier manifeste, consultez l’article [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="adbff-120">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="adbff-121">Résolution des erreurs de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="adbff-121">Troubleshoot user errors</span></span>

<span data-ttu-id="adbff-122">Pour plus d’informations sur la résolution des problèmes courants que les utilisateurs peuvent rencontrer avec votre complément Office, consultez [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="adbff-122">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
