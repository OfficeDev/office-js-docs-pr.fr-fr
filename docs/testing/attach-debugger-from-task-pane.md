---
title: Attacher un débogueur à partir du volet Office
description: Découvrez comment attacher un débogueur à partir du volet Office
ms.date: 09/09/2019
localization_priority: Normal
ms.openlocfilehash: 903ecfc577804ab052109d8a8f25c5a6eb799488
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611259"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Attacher un débogueur à partir du volet Office

Dans Office 2016 pour Windows, version 77xx.xxxx ou ultérieure, vous pouvez attacher le débogueur à partir du volet Office. Cette fonctionnalité attache directement le débogueur au processus Internet Explorer approprié pour vous. Vous pouvez attacher un débogueur quel que soit l’outil que vous utilisez (générateur de Yeoman, Visual Studio Code, Node.js, Angular ou autre). 

Pour lancer l’outil **Attacher le débogueur**, cliquez sur le coin supérieur droit du volet Office pour activer le menu **Caractéristique** (comme illustré dans le cercle rouge dans l’image suivante).   

> [!NOTE]
> - Actuellement, le seul débogueur pris en charge est [Visual Studio 2015](https://www.visualstudio.com/downloads/) avec la [mise à jour 3](https://msdn.microsoft.com/library/mt752379.aspx) ou une mise à jour ultérieure. Si Visual Studio n’est pas installé, la sélection de l’option **attacher le débogueur** n’entraîne aucune action.   
> - Vous ne pouvez déboguer JavaScript côté client qu’à l’aide de l’outil **Attacher le débogueur**. Pour déboguer du code côté serveur, comme avec un serveur Node.js, vous disposez de nombreuses options. Pour plus d’informations sur le débogage avec Visual Studio Code, reportez-vous à la rubrique sur le [débogage de Node.js dans VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Si vous n’utilisez pas Visual Studio Code, recherchez « déboguer Node.js » ou « déboguer {nom de serveur} ».

![Capture d’écran du menu Attacher le débogueur](../images/attach-debugger.png)

Sélectionnez **Attacher le débogueur**. Cette action ouvre la boîte de dialogue **Débogueur juste-à-temps Visual Studio**, comme illustré dans l’image suivante. 

![Capture d’écran de la boîte de dialogue Débogueur juste-à-temps Visual Studio](../images/visual-studio-debugger.png)

Dans Visual Studio, les fichiers de code s’affichent dans **l’Explorateur de solutions**.   Vous pouvez définir des points d’arrêt à la ligne de code que vous souhaitez déboguer dans Visual Studio.

> [!NOTE]
> Si vous ne voyez pas le menu Personnalité, vous pouvez déboguer votre complément à l’aide de Visual Studio. Vérifiez que votre complément de volet Office est ouvert dans Office, puis procédez comme suit :
>
> 1. Dans Visual Studio, choisissez **DÉBOGUER** > **Attacher au processus**.
> 2. Dans **Processus disponibles**, choisissez *soit* tous les processus `Iexplore.exe`disponibles, *soit* tous les processus `MicrosoftEdge*.exe` disponible selon que [votre complément utilise Internet Explorer ou Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md), puis cliquez sur le bouton **Joindre**.

Pour plus d’informations sur le débogage dans Visual Studio, consultez les rubriques suivantes :

-    Pour lancer et utiliser l’explorateur DOM dans Visual Studio, consultez le conseil 4 dans la section relative aux [conseils et astuces](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) du billet de blog sur la [création d’applications attrayantes pour Office à l’aide de nouveaux modèles de projet](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates).
-    Pour définir des points d’arrêt, consultez la rubrique [Utilisation des points d’arrêt](/visualstudio/debugger/using-breakpoints?view=vs-2015).
-    Pour utiliser F12, consultez la rubrique [Utilisation des outils de développement F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).
-   Pour utiliser les outils de développement Microsoft Edge, voir [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).

## <a name="see-also"></a>Voir aussi

- [Déboguer des compléments Office dans Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md)
- [Publier votre complément Office](../publish/publish.md)
