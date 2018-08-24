---
title: Test et débogage de compléments Office
description: ''
ms.date: 11/24/2017
ms.openlocfilehash: f645482fa92faad2e28484fa4b878bd35c0a27b6
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925261"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="59f8b-102">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="59f8b-102">Test and debug Office Add-ins</span></span>

<span data-ttu-id="59f8b-103">Cette section contient des recommandations sur les tests, le débogage et la résolution des problèmes avec les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="59f8b-103">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="59f8b-104">Chargement de version test d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="59f8b-104">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="59f8b-105">Vous pouvez utiliser le chargement de version test pour installer un complément Office sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="59f8b-105">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog.</span></span> <span data-ttu-id="59f8b-106">La procédure de chargement de version test d’un complément varie en fonction de la plateforme et, dans certains cas, du produit.</span><span class="sxs-lookup"><span data-stu-id="59f8b-106">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="59f8b-107">Les articles suivants décrivent comment charger une version test de compléments Office sur une plateforme spécifique ou dans un produit spécifique :</span><span class="sxs-lookup"><span data-stu-id="59f8b-107">The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="59f8b-108">Chargement de version test des compléments Office sur Windows</span><span class="sxs-lookup"><span data-stu-id="59f8b-108">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="59f8b-109">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="59f8b-109">Sideload Office Add-ins in Office Online</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="59f8b-110">Chargement de version test de compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="59f8b-110">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- <span data-ttu-id="59f8b-111">
  [Chargement de version test des compléments Outlook](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)</span><span class="sxs-lookup"><span data-stu-id="59f8b-111">[Sideload Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)</span></span>

## <a name="debug-an-office-add-in"></a><span data-ttu-id="59f8b-112">Débogage d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="59f8b-112">Debug an Office Add-in</span></span>

<span data-ttu-id="59f8b-113">La procédure pour le débogage d’un complément Office varie également selon la plateforme.</span><span class="sxs-lookup"><span data-stu-id="59f8b-113">The procedure for debugging an Office Add-in varies by platform as well.</span></span> <span data-ttu-id="59f8b-114">Chacun des articles suivants décrit comment déboguer des compléments Office sur une plateforme spécifique :</span><span class="sxs-lookup"><span data-stu-id="59f8b-114">Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="59f8b-115">Attacher un débogueur à partir du volet Office (sur Windows)</span><span class="sxs-lookup"><span data-stu-id="59f8b-115">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="59f8b-116">Débogage de compléments avec les outils de développement F12 sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="59f8b-116">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="59f8b-117">Débogage de compléments dans Office Online</span><span class="sxs-lookup"><span data-stu-id="59f8b-117">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="59f8b-118">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="59f8b-118">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="59f8b-119">Validation d’un manifeste de complément Office</span><span class="sxs-lookup"><span data-stu-id="59f8b-119">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="59f8b-120">Pour savoir comment valider le fichier manifeste qui décrit votre complément Office et résoudre des problèmes avec le fichier manifeste, consultez l’article [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="59f8b-120">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="59f8b-121">Résolution des erreurs de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="59f8b-121">Troubleshoot user errors</span></span>

<span data-ttu-id="59f8b-122">Pour plus d’informations sur la résolution des problèmes courants que les utilisateurs peuvent rencontrer avec votre complément Office, consultez [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="59f8b-122">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>