---
title: Utiliser la boîte de dialogue Office pour lire une vidéo
description: Découvrez comment ouvrir et lire une vidéo dans la boîte de dialogue Office de lecture
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 2519b2f105503a0479eee07d885a1543f5455343
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349881"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="926df-103">Utiliser la boîte Office dialogue pour afficher une vidéo</span><span class="sxs-lookup"><span data-stu-id="926df-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="926df-104">Cet article explique comment lire une vidéo dans une boîte de dialogue Office de l’article.</span><span class="sxs-lookup"><span data-stu-id="926df-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="926df-105">Cet article suppose que vous connaissez les principes de base de l’utilisation de la boîte de dialogue Office, comme décrit dans [l’API](dialog-api-in-office-add-ins.md)de boîte de dialogue Office dans vos Office.</span><span class="sxs-lookup"><span data-stu-id="926df-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="926df-106">Pour lire une vidéo dans une boîte de dialogue avec l’API Office boîte de dialogue, suivez les étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="926df-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="926df-107">Créez une page contenant un iframe et aucun autre contenu.</span><span class="sxs-lookup"><span data-stu-id="926df-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="926df-108">La page doit se trouver dans le même domaine que la page hôte.</span><span class="sxs-lookup"><span data-stu-id="926df-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="926df-109">Pour un rappel de ce qu’est une page hôte, voir Ouvrir une boîte de [dialogue à partir d’une page hôte.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)</span><span class="sxs-lookup"><span data-stu-id="926df-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="926df-110">Dans `src` l’attribut de l’iframe, pointer vers l’URL d’une vidéo en ligne.</span><span class="sxs-lookup"><span data-stu-id="926df-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="926df-111">Le protocole de l’URL de la vidéo doit être HTTPS.</span><span class="sxs-lookup"><span data-stu-id="926df-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="926df-112">Dans cet article, nous appellerons cette page « video.dialogbox.html ».</span><span class="sxs-lookup"><span data-stu-id="926df-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="926df-113">Voici un exemple de marques de révision.</span><span class="sxs-lookup"><span data-stu-id="926df-113">The following is an example of the markup.</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="926df-114">Utilisez un appel de `displayDialogAsync` dans la page hôte pour ouvrir video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="926df-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="926df-115">Si votre complément a besoin de savoir quand l’utilisateur ferme la boîte de dialogue, inscrivez un gestionnaire pour l’événement `DialogEventReceived` et gérez l’événement 12006.</span><span class="sxs-lookup"><span data-stu-id="926df-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="926df-116">Pour plus d’informations, voir [Erreurs et événements dans la boîte Office dialogue .](dialog-handle-errors-events.md)</span><span class="sxs-lookup"><span data-stu-id="926df-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="926df-117">Pour obtenir un exemple de vidéo en cours de lecture dans une boîte de dialogue, voir le modèle de conception [de placemat vidéo.](../design/first-run-experience-patterns.md#video-placemat)</span><span class="sxs-lookup"><span data-stu-id="926df-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Capture d’écran montrant une vidéo en cours de lecture dans une boîte de dialogue de Excel.](../images/video-placemats-dialog-open.png)
