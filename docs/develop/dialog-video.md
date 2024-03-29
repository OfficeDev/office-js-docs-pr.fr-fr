---
title: Utiliser la boîte de dialogue Office pour lire une vidéo
description: Découvrez comment ouvrir et lire une vidéo dans la boîte Office dialogue.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: a9f222f52d1ee22a4b0b84eb62ea24e6e48e63d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743509"
---
# <a name="use-the-office-dialog-box-to-show-a-video"></a>Utiliser la boîte Office dialogue pour afficher une vidéo

Cet article explique comment lire une vidéo dans une boîte de dialogue Office de l’objet Add-in.

> [!NOTE]
> Cet article suppose que vous connaissez les principes de base de l’utilisation de la boîte de dialogue Office, comme décrit dans [l’API](dialog-api-in-office-add-ins.md) de boîte de dialogue Office dans vos Office.

Pour lire une vidéo dans une boîte de dialogue avec l’API Office boîte de dialogue, suivez ces étapes.

1. Créez une page contenant un iframe et aucun autre contenu. La page doit se trouver dans le même domaine que la page hôte. Pour un rappel de ce qu’est une page hôte, voir [Ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page). Dans l’attribut `src` de l’iframe, pointer vers l’URL d’une vidéo en ligne. Le protocole de l’URL de la vidéo doit être HTTPS. Dans cet article, nous appellerons cette page « video.dialogbox.html ». Voici un exemple de marques de révision.

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Utilisez un appel de `displayDialogAsync` dans la page hôte pour ouvrir video.dialogbox.html.
3. Si votre complément a besoin de savoir quand l’utilisateur ferme la boîte de dialogue, inscrivez un gestionnaire pour l’événement `DialogEventReceived` et gérez l’événement 12006. Pour plus d’informations, voir [Erreurs et événements dans la boîte Office dialogue.](dialog-handle-errors-events.md)

Pour obtenir un exemple de vidéo en cours de lecture dans une boîte de dialogue, voir le modèle de conception [de placemat vidéo](../design/first-run-experience-patterns.md#video-placemat).

![Capture d’écran montrant une vidéo en cours de lecture dans une boîte de dialogue de Excel.](../images/video-placemats-dialog-open.png)
