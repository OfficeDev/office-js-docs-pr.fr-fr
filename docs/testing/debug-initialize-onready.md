---
title: Déboguer les fonctions initialiser et onReady
description: Découvrez comment déboguer les fonctions Office.initialize et Office.onReady.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d26c0f2686477f74fbe65550b7d35cc69e7be71
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958845"
---
# <a name="debug-the-initialize-and-onready-functions"></a>Déboguer les fonctions initialiser et onReady

> [!NOTE]
> Cet article part du principe que vous connaissez bien [Initialiser votre complément Office](../develop/initialize-add-in.md).

Le paradoxe du débogage des fonctions [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) et [Office.onReady](/javascript/api/office#office-office-onready-function(1)) est qu’un débogueur ne peut s’attacher qu’à un processus en cours d’exécution, mais ces fonctions s’exécutent immédiatement au démarrage du processus d’exécution du complément, avant qu’un débogueur puisse s’attacher. Dans la plupart des cas, le redémarrage du complément après l’attachement d’un débogueur n’est pas utile, car le redémarrage du complément ferme le processus d’exécution d’origine *et le débogueur attaché* et démarre un nouveau processus sans débogueur attaché.

Heureusement, il existe une exception. Vous pouvez déboguer ces fonctions à l’aide de Office sur le Web, en procédant comme suit.

1. Chargez et exécutez le complément dans Office sur le Web. Pour ce faire, vous devez généralement ouvrir le volet Office d’un complément ou exécuter une [commande de fonction](../design/add-in-commands.md#types-of-add-in-commands). *Le complément s’exécute dans le processus de navigateur global, et non dans un processus distinct comme dans Office de bureau.*
1. Ouvrez les outils de développement du navigateur. Pour ce faire, appuyez généralement sur F12. Le débogueur dans les outils s’attache au processus du navigateur.
1. Appliquez des points d’arrêt en fonction des besoins au code de la ou `Office.onReady` de la `Office.initialize` fonction.
1. *Relancez le volet Office du complément ou la commande de fonction* comme vous l’avez fait à l’étape 1. Cette action ne ferme *pas* le processus de navigateur ou le débogueur. La ou `Office.onReady` la `Office.initialize` fonction s’exécute à nouveau et le traitement s’arrête sur vos points d’arrêt.

> [!TIP]
> Pour plus d’informations, consultez [Les compléments de débogage dans Office sur le Web](debug-add-ins-in-office-online.md).
