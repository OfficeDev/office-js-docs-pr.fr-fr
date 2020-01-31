---
title: Vider le cache Office
description: Découvrez comment effacer le cache Office sur votre ordinateur.
ms.date: 01/21/2020
localization_priority: Priority
ms.openlocfilehash: 68e5c022671844ee44bf8ca8ac00bc5af6564bad
ms.sourcegitcommit: 43166612e9b4bf7a73312a572663c8696353dbc6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/29/2020
ms.locfileid: "41580962"
---
# <a name="clear-the-office-cache"></a>Vider le cache Office

Vous pouvez supprimer un complément que vous avez précédemment chargé sur Windows, Mac ou iOS en vidant le cache Office sur votre ordinateur. 

En outre, si vous apportez des modifications au manifeste de votre complément (par exemple, vous mettez à jour le nom des fichiers d’icônes ou de texte de commandes du complément), videz le cache Office, puis rechargez le complément à l’aide d’un manifeste mis à jour. Cette action permettra à Office d’afficher le complément tel que décrit par le manifeste mis à jour.

## <a name="clear-the-office-cache-on-windows"></a>Vider le cache Office sur Windows

### <a name="excel-word-and-powerpoint"></a>Excel, Word et PowerPoint 

Pour vider le cache Office sur Windows pour Excel, Word et PowerPoint, supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

### <a name="outlook-windows-10"></a>Outlook (Windows 10)

Pour vider le cache Outlook dans Windows 10 lorsque le complément s’exécute dans Microsoft Edge, vous pouvez utiliser Microsoft Edge DevTools.

> [!TIP]
> Si vous souhaitez que le complément sideloaded reflète les modifications récentes apportées à ses fichiers HTML ou JavaScript, il n’est pas nécessaire que vous utilisiez les étapes suivantes pour vider le cache. Il vous suffit, au lieu de cela, d’insérer le focus dans le volet de tâches du complément (en cliquant n’importe où dans le volet), puis d’appuyer sur **F5** pour recharger le complément. 

> [!NOTE]
> Pour effacer le cache Outlook en suivant les étapes ci-dessous, votre complément doit avoir un volet de tâches. Si vous avez un complément UI-less, par exemple un complément qui utilise la fonctionnalité [on-send](/outlook/add-ins/outlook-on-send-addins), vous devez ajouter un volet de tâches à votre complément qui utilise le même domaine pour [SourceLocation](../reference/manifest/sourcelocation.md), avant de pouvoir utiliser les étapes suivantes pour vider le cache.

1. Installez [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Ouvrez votre complément dans Outlook.

3. Exécutez Microsoft Edge DevTools.

4. Ouvrez l’onglet **Local** dans Microsoft Edge DevTools. Votre complément est répertorié par son nom.

5. Sélectionnez le nom du complément pour joindre le débogueur à votre complément. Une nouvelle fenêtre Microsoft Edge DevTools s’ouvre lorsque le débogueur s'attache à votre complément.

6. Sous l’onglet **Réseau** de la nouvelle fenêtre, sélectionnez le bouton **Vider le cache**.

    ![Capture d’écran Microsoft Edge DevTools avec le bouton Vider le cache mis en évidence](../images/edge-devtools-clear-cache.png)

7. Si l’exécution de ces étapes ne produit pas le résultat escompté, vous pouvez également sélectionner le bouton **Toujours actualiser à partir du serveur**.

    ![Capture d’écran Microsoft Edge DevTools avec le bouton Toujours actualiser à partir du serveur mis en évidence](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Vider le cache Office sur Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a>Vider le cache Office sur iOS

Pour vider le cache Office sur iOS, appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.

## <a name="see-also"></a>Voir aussi

- [Débogage des compléments Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)

