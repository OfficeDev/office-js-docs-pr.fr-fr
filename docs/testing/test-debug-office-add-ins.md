---
title: Test et débogage de compléments Office
description: Découvrez comment tester et déboguer votre Complément Office.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: b620d351c4d584a933431bc6d23a88f1a1de0962
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719741"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="e9418-103">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="e9418-103">Test and debug Office Add-ins</span></span>

<span data-ttu-id="e9418-104">Cette section contient des recommandations sur les tests, le débogage et la résolution des problèmes avec les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="e9418-104">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="e9418-105">Chargement de version test d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="e9418-105">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="e9418-p101">Vous pouvez utiliser le chargement de version test pour installer un complément Office sans avoir à le placer au préalable dans un catalogue de compléments. La procédure de chargement de version test d’un complément varie en fonction de la plateforme et, dans certains cas, du produit. Les articles suivants décrivent comment charger une version test de compléments Office sur une plateforme spécifique ou dans un produit spécifique :</span><span class="sxs-lookup"><span data-stu-id="e9418-p101">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="e9418-109">Chargement de version test des compléments Office sur Windows</span><span class="sxs-lookup"><span data-stu-id="e9418-109">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="e9418-110">Chargement de version test des compléments Office dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e9418-110">Sideload Office Add-ins in Office on the web</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="e9418-111">Chargement de version test de compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="e9418-111">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="e9418-112">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="e9418-112">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="e9418-113">Débogage d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="e9418-113">Debug an Office Add-in</span></span>

<span data-ttu-id="e9418-p102">La procédure pour le débogage d’un complément Office varie également selon la plateforme. Chacun des articles suivants décrit comment déboguer des compléments Office sur une plateforme spécifique :</span><span class="sxs-lookup"><span data-stu-id="e9418-p102">The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="e9418-116">Attacher un débogueur à partir du volet Office (sur Windows)</span><span class="sxs-lookup"><span data-stu-id="e9418-116">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="e9418-117">Débogage de compléments avec les outils de développement F12 sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="e9418-117">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="e9418-118">Débogage de compléments dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e9418-118">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="e9418-119">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="e9418-119">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="e9418-120">Validation d’un manifeste de complément Office</span><span class="sxs-lookup"><span data-stu-id="e9418-120">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="e9418-121">Pour savoir comment valider le fichier manifeste qui décrit votre complément Office et résoudre des problèmes avec le fichier manifeste, consultez l’article [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="e9418-121">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="e9418-122">Résolution des erreurs de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="e9418-122">Troubleshoot user errors</span></span>

<span data-ttu-id="e9418-123">Pour plus d’informations sur la résolution des problèmes courants que les utilisateurs peuvent rencontrer avec votre complément Office, consultez [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="e9418-123">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
