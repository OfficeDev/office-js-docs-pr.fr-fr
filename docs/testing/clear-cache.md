---
title: Vider le cache Office
description: Découvrez comment effacer le cache Office sur votre ordinateur.
ms.date: 03/11/2022
ms.localizationpriority: high
ms.openlocfilehash: 04a329e9e7289f90b02b9307c67eef2818191221
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711230"
---
# <a name="clear-the-office-cache"></a>Vider le cache Office

Pour supprimer un complément que vous avez précédemment chargé sur Windows, Mac ou iOS, vous devez effacer le cache Office sur votre ordinateur.

En outre, si vous apportez des modifications au manifeste de votre complément (par exemple, vous mettez à jour le nom des fichiers d’icônes ou de texte de commandes du complément), videz le cache Office, puis rechargez le complément à l’aide d’un manifeste mis à jour. Cette action permettra à Office d’afficher le complément tel que décrit par le manifeste mis à jour.

> [!NOTE]
> Pour supprimer un complément chargé de manière indépendante d’Excel, OneNote, PowerPoint ou Word sur le web, consultez [Charger une version test des compléments Office dans Office sur le web à des fins de test : Supprimer un complément chargé de manière indépendante](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in).

## <a name="clear-the-office-cache-on-windows"></a>Vider le cache Office sur Windows

Il existe trois façons d’effacer le cache Office sur un ordinateur Windows : automatiquement, manuellement et à l’aide des outils de développement Microsoft Edge. Les méthodes sont décrites dans les sous-sections suivantes.

### <a name="automatically"></a>Automatiquement

Cette méthode est recommandée pour les ordinateurs de développement de complément. Si votre version d’Office sur Windows est 2108 ou ultérieure, les étapes suivantes configurent le cache Office pour qu’il soit effacé lors de la prochaine réouverture d’Office.

> [!NOTE]
> La méthode automatique n’est pas prise en charge pour Outlook.

1. À partir du ruban de n’importe quel hôte Office à l’exception d’Outlook, accédez à **Fichier** > **Options** > **Centre de gestion de la confidentialité** > **Paramètres du Centre de gestion de la confidentialité** > **Catalogues des compléments approuvés**.
1. Sélectionnez la case **Au prochain démarrage d’Office, effacer le cache de tous les compléments web précédemment démarrés**.

### <a name="manually"></a>Manuellement

La méthode manuelle pour Excel, Word et PowerPoint est différente de celle pour Outlook.

#### <a name="manually-clear-the-cache-in-excel-word-and-powerpoint"></a>Effacer manuellement le cache dans Excel, Word et PowerPoint

Pour supprimer tous les compléments chargés indépendamment d’Excel, de Word et de PowerPoint, supprimez le contenu du dossier suivant.

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Si le dossier suivant existe, supprimez également son contenu.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

#### <a name="manually-clear-the-cache-in-outlook"></a>Effacer manuellement le cache dans Outlook

Pour supprimer un complément chargé indépendamment d’Outlook, suivez la procédure décrite dans [Charger indépendamment des compléments Outlook à des fins de test](../outlook/sideload-outlook-add-ins-for-testing.md) pour rechercher le complément dans la section **Compléments personnalisés** de la boîte de dialogue qui répertorie les compléments installés. Sélectionnez les points de suspension (`...`) du complément, puis sélectionnez **Supprimer** pour supprimer ce complément spécifique. Si la suppression de ce complément ne fonctionne pas, supprimez le contenu du dossier `Wef`comme indiqué précédemment pour Excel, Word et PowerPoint.

### <a name="using-the-microsoft-edge-developer-tools"></a>Utilisation des outils de développement Microsoft Edge

Pour effacer le cache Office sur Windows 10 lorsque le complément s’exécute dans Microsoft Edge, vous pouvez utiliser Outils de développement Microsoft Edge.

> [!TIP]
> Si vous souhaitez que le complément chargé indépendamment reflète les modifications récentes apportées à ses fichiers sources HTML ou JavaScript, vous n’avez normalement pas besoin de vider le cache. Il vous suffit, au lieu de cela, d’insérer le focus dans le volet de tâches du complément (en cliquant n’importe où dans le volet), puis d’appuyer sur **Ctrl+F5** pour recharger le complément.

> [!NOTE]
> Pour effacer le cache d'Office à l'aide des étapes suivantes, votre module complémentaire doit disposer d'un volet de tâches. Si vous avez un complément UI-less, par exemple un complément qui utilise la fonctionnalité [on-send](../outlook/outlook-on-send-addins.md), vous devez ajouter un volet de tâches à votre complément qui utilise le même domaine pour [SourceLocation](../reference/manifest/sourcelocation.md), avant de pouvoir utiliser les étapes suivantes pour vider le cache.

1. Installez [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).

2. Ouvrez votre complément dans le client Office.

3. Exécutez Microsoft Edge DevTools.

4. Ouvrez l’onglet **Local** dans Microsoft Edge DevTools. Votre complément est répertorié par son nom.

5. Sélectionnez le nom du complément pour joindre le débogueur à votre complément. Une nouvelle fenêtre Microsoft Edge DevTools s’ouvre lorsque le débogueur s'attache à votre complément.

6. Sous l’onglet **Réseau** de la nouvelle fenêtre, sélectionnez le bouton **Vider le cache**.

    ![Capture d’écran Microsoft Edge DevTools avec le bouton Vider le cache mis en évidence.](../images/edge-devtools-clear-cache.png)

7. Si l’exécution de ces étapes ne produit pas le résultat escompté, vous pouvez également sélectionner le bouton **Toujours actualiser à partir du serveur**.

    ![Capture d’écran Microsoft Edge DevTools avec le bouton Toujours actualiser à partir du serveur mis en évidence.](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Vider le cache Office sur Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>Vider le cache Office sur iOS

Pour vider le cache Office sur iOS, appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également réinstaller Office.

## <a name="see-also"></a>Voir aussi

- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](troubleshoot-development-errors.md)
- [Déboguer des compléments à l’aide des outils de développement pour Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Déboguer des compléments à l’aide des outils de développement pour la version héritée Edge](debug-add-ins-using-devtools-edge-legacy.md)
- [Déboguer des compléments à l’aide des Outils de développement dans Microsoft Edge (basés sur Chromium)](debug-add-ins-using-devtools-edge-chromium.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Chargement de la version test des compléments Office](sideload-office-add-ins-for-testing.md)
- [Manifeste XML des compléments Office](../develop/add-in-manifests.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
