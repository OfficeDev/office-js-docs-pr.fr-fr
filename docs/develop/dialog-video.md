---
title: Utiliser la boîte de dialogue Office pour lire une vidéo
description: Obtenir des informations sur l’ouverture et la lecture d’une vidéo dans la boîte de dialogue Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: f0d524996b105061b8e5d1b584d8b3e0d44eec7c
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131772"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>Afficher une vidéo à l’aide de la boîte de dialogue Office

Cet article explique comment lire une vidéo dans une boîte de dialogue de complément Office.

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec les notions de base de l’utilisation de la boîte de dialogue Office, comme décrit dans la rubrique [utiliser l’API de boîte de dialogue Office dans vos compléments Office](dialog-api-in-office-add-ins.md).

Pour lire une vidéo dans une boîte de dialogue avec l’API de boîte de dialogue Office, procédez comme suit :

1. Créez une page contenant un IFRAME et aucun autre contenu. La page doit se trouver dans le même domaine que la page hôte. Pour un rappel de ce qu’est une page hôte, voir [ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page). Dans l' `src` attribut de l’IFRAME, pointez sur l’URL d’une vidéo en ligne. Le protocole de l’URL de la vidéo doit être HTTPS. Dans cet article, nous allons appeler cette page « video.dialogbox.html ». Voici un exemple de marques de révision :

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Utilisez un appel de `displayDialogAsync` dans la page hôte pour ouvrir video.dialogbox.html.
3. Si votre complément a besoin de savoir quand l’utilisateur ferme la boîte de dialogue, inscrivez un gestionnaire pour l’événement `DialogEventReceived` et gérez l’événement 12006. Pour plus d’informations, consultez [la rubrique Erreurs et événements dans la boîte de dialogue Office](dialog-handle-errors-events.md).

Pour voir un exemple de lecture vidéo dans une boîte de dialogue, consultez le modèle de conception de la vidéo de [Présentation vidéo](../design/first-run-experience-patterns.md#video-placemat).

![Capture d’écran illustrant la lecture d’une vidéo dans une boîte de dialogue de complément devant Excel](../images/video-placemats-dialog-open.png)
