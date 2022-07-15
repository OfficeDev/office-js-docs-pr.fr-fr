---
title: Déboguer une commande de fonction avec un runtime non partagé
description: Découvrez comment déboguer des commandes de fonction.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 943d7ed8ccfedd961eac3fe941c8ef357964ed37
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797704"
---
# <a name="debug-a-function-command-with-a-non-shared-runtime"></a>Déboguer une commande de fonction avec un runtime non partagé

> [!IMPORTANT]
> Si votre complément est [configuré pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md), vous déboguez le code derrière la commande de fonction comme vous le feriez pour le code situé derrière un volet Office. Consultez [Déboguer des compléments Office](debug-add-ins-overview.md) et notez qu’une commande de fonction dans un complément avec un runtime partagé *n’est pas* un cas particulier comme décrit dans cet article. 

> [!NOTE]
> Cet article suppose que vous êtes familiarisé avec [les commandes de fonction](../design/add-in-commands.md#types-of-add-in-commands).

Les commandes de fonction n’ayant pas d’interface utilisateur, un débogueur ne peut pas être attaché au processus dans lequel la fonction s’exécute sur office de bureau. (Les compléments Outlook en cours de développement sur Windows sont une exception à cette règle. Consultez [les commandes de fonction de débogage dans les compléments Outlook sur Windows](#debug-function-commands-in-outlook-add-ins-on-windows) plus loin dans cet article.) Par conséquent, les commandes de fonction, dans les compléments avec un runtime non partagé, doivent être déboguées sur Office sur le Web où la fonction s’exécute dans le processus de navigateur global. Suivez les étapes ci-dessous.

1. Chargez le complément dans Office sur le Web, puis sélectionnez le bouton ou l’élément de menu qui exécute la commande de fonction. Cela est nécessaire pour charger le fichier de code de la commande de fonction. 
1. Ouvrez les outils de développement du navigateur. Pour ce faire, appuyez généralement sur F12. Le débogueur dans les outils s’attache au processus du navigateur.
1. Appliquez des points d’arrêt au code en fonction des besoins de la commande de fonction.
1. Réexécutez la commande de fonction. Le processus s’arrête sur vos points d’arrêt. 

> [!TIP]
> Pour plus d’informations, consultez [Les compléments de débogage dans Office sur le Web](debug-add-ins-in-office-online.md).

## <a name="debug-function-commands-in-outlook-add-ins-on-windows"></a>Commandes de fonction de débogage dans les compléments Outlook sur Windows

Si votre ordinateur de développement est Windows, vous pouvez déboguer une commande de fonction sur le bureau Outlook. Consultez [les commandes de fonction de débogage dans les compléments Outlook](../outlook/debug-ui-less.md).