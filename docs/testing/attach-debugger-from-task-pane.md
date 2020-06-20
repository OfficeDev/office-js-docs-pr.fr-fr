---
title: Attacher un débogueur à partir du volet Office
description: Découvrez comment attacher un débogueur à partir du volet Office
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 53cfce211241dbdf3d16e8a126e059a2f2db3f23
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810841"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Attacher un débogueur à partir du volet Office

In Office 2016 on Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, Node.js, Angular, or another tool. 

Pour lancer l’outil **Attacher le débogueur**, cliquez sur le coin supérieur droit du volet Office pour activer le menu **Caractéristique** (comme illustré dans le cercle rouge dans l’image suivante).   

> [!NOTE]
> - Actuellement, le seul débogueur pris en charge est [Visual Studio 2015](https://www.visualstudio.com/downloads/) avec la [mise à jour 3](https://msdn.microsoft.com/library/mt752379.aspx) ou une mise à jour ultérieure. Si Visual Studio n’est pas installé, la sélection de l’option **attacher le débogueur** n’entraîne aucune action.   
> - You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".

![Capture d’écran du menu Attacher le débogueur](../images/attach-debugger.png)

Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image. 

![Capture d’écran de la boîte de dialogue Débogueur juste-à-temps Visual Studio](../images/visual-studio-debugger.png)

In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.

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
- [Extension du débogueur de complément Microsoft Office pour Visual Studio code](debug-with-vs-extension.md)