---
title: Attacher un débogueur à partir du volet Office
description: Découvrez comment attacher un débogger à partir du volet Des tâches
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 3efe02f2683990a8f4d802bff5040ba9e007c5c91574bba274f4c26b9a5b8683
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086674"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Attacher un débogueur à partir du volet Office

Dans Office 2016 pour Windows, version 77xx.xxxx ou ultérieure, vous pouvez attacher le débogueur à partir du volet Office. Cette fonctionnalité attache directement le débogueur au processus Internet Explorer approprié pour vous. Vous pouvez attacher un débogueur quel que soit l’outil que vous utilisez (générateur de Yeoman, Visual Studio Code, Node.js, Angular ou autre).

Pour lancer l’outil **Attacher le débogueur**, cliquez sur le coin supérieur droit du volet Office pour activer le menu **Caractéristique** (comme illustré dans le cercle rouge dans l’image suivante).

> [!NOTE]
> - Actuellement, le seul outil de débogger pris en charge [est Visual Studio 2015](https://www.visualstudio.com/downloads/) avec la mise à jour [3](/previous-versions/mt752379(v=vs.140)) ou ultérieure. Si vous n’avez pas installé Visual Studio, la  sélection de l’option Attacher le débogger n’entraîne aucune action.
> - Vous ne pouvez déboguer JavaScript côté client qu’à l’aide de l’outil **Attacher le débogueur**. Pour déboguer du code côté serveur, comme avec un serveur Node.js, vous disposez de nombreuses options. Pour plus d’informations sur le débogage avec Visual Studio Code, reportez-vous à la rubrique sur le [débogage de Node.js dans VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Si vous n’utilisez pas Visual Studio Code, recherchez « déboguer Node.js » ou « déboguer {nom de serveur} ».

![Capture d’écran du menu Attacher le débogger.](../images/attach-debugger.png)

Sélectionnez **Attacher le débogueur**. Cette action ouvre la boîte de dialogue **Débogueur juste-à-temps Visual Studio**, comme illustré dans l’image suivante.

![Capture d’Visual Studio boîte de dialogue Débogger JIT.](../images/visual-studio-debugger.png)

Dans Visual Studio, les fichiers de code s’affichent dans **l’Explorateur de solutions**.   Vous pouvez définir des points d’arrêt à la ligne de code que vous souhaitez déboguer dans Visual Studio.

> [!NOTE]
> Si vous ne voyez pas le menu Personnalité, vous pouvez déboguer votre complément à l’aide de Visual Studio. Assurez-vous que votre add-in du volet Des tâches est ouvert dans Office, puis suivez ces étapes.
>
> 1. Dans Visual Studio, choisissez **DÉBOGUER** > **Attacher au processus**.
> 2. Dans **Processus disponibles**, choisissez *soit* tous les processus `Iexplore.exe`disponibles, *soit* tous les processus `MicrosoftEdge*.exe` disponible selon que [votre complément utilise Internet Explorer ou Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md), puis cliquez sur le bouton **Joindre**.

Pour plus d’informations sur le débogage dans Visual Studio, consultez les rubriques suivantes :

- Pour lancer et utiliser l’explorateur DOM dans Visual Studio, consultez le conseil 4 dans la section relative aux [conseils et astuces](/archive/blogs/officeapps/building-great-looking-apps-for-office-using-the-new-project-templates#tips_tricks) du billet de blog sur la [création d’applications attrayantes pour Office à l’aide de nouveaux modèles de projet](/archive/blogs/officeapps/building-great-looking-apps-for-office-using-the-new-project-templates).
- Pour définir des points d’arrêt, consultez la rubrique [Utilisation des points d’arrêt](/visualstudio/debugger/using-breakpoints?view=vs-2015&preserve-view=true).
- Pour utiliser F12, consultez la rubrique [Utilisation des outils de développement F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).
- Pour utiliser les outils de développement Microsoft Edge, voir [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).

## <a name="see-also"></a>Voir aussi

- [Déboguer des compléments Office dans Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md)
- [Publier votre complément Office](../publish/publish.md)
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)