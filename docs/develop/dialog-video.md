---
title: Utiliser la boîte de dialogue Office pour lire une vidéo
description: Obtenir des informations sur l’ouverture et la lecture d’une vidéo dans la boîte de dialogue Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 9c65dfb9c0cf1adbc827be25b655e380dc39e2d2
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596528"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a><span data-ttu-id="b71d0-103">Afficher une vidéo à l’aide de la boîte de dialogue Office</span><span class="sxs-lookup"><span data-stu-id="b71d0-103">Use the Office dialog box to show a video</span></span>

<span data-ttu-id="b71d0-104">Cet article explique comment lire une vidéo dans une boîte de dialogue de complément Office.</span><span class="sxs-lookup"><span data-stu-id="b71d0-104">This article explains how to play a video in an Office Add-in dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="b71d0-105">Cet article suppose que vous êtes familiarisé avec les notions de base de l’utilisation de la boîte de dialogue Office, comme décrit dans la rubrique [utiliser l’API de boîte de dialogue Office dans vos compléments Office](dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b71d0-105">This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>

<span data-ttu-id="b71d0-106">Pour lire une vidéo dans une boîte de dialogue avec l’API de boîte de dialogue Office, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="b71d0-106">To play a video in a dialog box with the Office dialog API, follow these steps:</span></span>

1. <span data-ttu-id="b71d0-107">Créez une page contenant un IFRAME et aucun autre contenu.</span><span class="sxs-lookup"><span data-stu-id="b71d0-107">Create a page containing an iframe and no other content.</span></span> <span data-ttu-id="b71d0-108">La page doit se trouver dans le même domaine que la page hôte.</span><span class="sxs-lookup"><span data-stu-id="b71d0-108">The page must be in the same domain as the host page.</span></span> <span data-ttu-id="b71d0-109">Pour un rappel de ce qu’est une page hôte, voir [ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="b71d0-109">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span> <span data-ttu-id="b71d0-110">Dans l' `src` attribut de l’IFRAME, pointez sur l’URL d’une vidéo en ligne.</span><span class="sxs-lookup"><span data-stu-id="b71d0-110">In the `src` attribute of the iframe, point to the URL of an online video.</span></span> <span data-ttu-id="b71d0-111">Le protocole de l’URL de la vidéo doit être HTTPS.</span><span class="sxs-lookup"><span data-stu-id="b71d0-111">The protocol of the video's URL must be HTTPS.</span></span> <span data-ttu-id="b71d0-112">Dans cet article, nous allons appeler cette page « Video. DialogBox. html ».</span><span class="sxs-lookup"><span data-stu-id="b71d0-112">In this article, we'll call this page "video.dialogbox.html".</span></span> <span data-ttu-id="b71d0-113">Voici un exemple de marques de révision :</span><span class="sxs-lookup"><span data-stu-id="b71d0-113">The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. <span data-ttu-id="b71d0-114">Utilisez un appel de `displayDialogAsync` dans la page hôte pour ouvrir video.dialogbox.html.</span><span class="sxs-lookup"><span data-stu-id="b71d0-114">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
3. <span data-ttu-id="b71d0-115">Si votre complément a besoin de savoir quand l’utilisateur ferme la boîte de dialogue, inscrivez un gestionnaire pour l’événement `DialogEventReceived` et gérez l’événement 12006.</span><span class="sxs-lookup"><span data-stu-id="b71d0-115">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event.</span></span> <span data-ttu-id="b71d0-116">Pour plus d’informations, consultez [la rubrique Erreurs et événements dans la boîte de dialogue Office](dialog-handle-errors-events.md).</span><span class="sxs-lookup"><span data-stu-id="b71d0-116">For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

<span data-ttu-id="b71d0-117">Pour voir un exemple de lecture vidéo dans une boîte de dialogue, consultez le modèle de conception de la vidéo de [Présentation vidéo](../design/first-run-experience-patterns.md#video-placemat).</span><span class="sxs-lookup"><span data-stu-id="b71d0-117">For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).</span></span>

![Capture d’écran d’une lecture vidéo dans une boîte de dialogue de complément](../images/video-placemats-dialog-open.png)
