---
title: Déboguer votre complément sans interface utilisateur Outlook
description: Découvrez comment déboguer votre complément sans interface utilisateur Outlook.
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 33aa36f86b7a163e650a23296b4c35aca7cb5492
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650708"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>Déboguer votre complément sans interface utilisateur Outlook

Cet article explique comment utiliser l’extension de débogueur de complément Office dans Visual Studio Code pour déboguer [des compléments sans interface utilisateur Outlook](add-in-commands-for-outlook.md#executing-a-javascript-function). Les actions de complément sans interface utilisateur sont lancées via un bouton de commande de complément dans le ruban. Pour plus d’informations sur les commandes de complément, consultez [les commandes de complément pour Outlook](add-in-commands-for-outlook.md).

Cet article part du principe que vous disposez déjà d’un projet de complément que vous souhaitez déboguer. Pour créer un complément sans interface utilisateur pour effectuer le débogage, suivez les étapes décrites dans le [didacticiel : Créer un complément de composition de message Outlook complément](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Marquer votre complément pour le débogage

Si vous avez utilisé le [générateur Yeoman pour Office compléments](../develop/yeoman-generator-overview.md) pour créer votre projet de complément, passez à la section [Configurer et exécutez le débogueur](#configure-and-run-the-debugger) plus loin dans cet article. Lorsque vous exécutez `npm start` pour générer votre complément et démarrer le serveur local, la commande définit également la `UseDirectDebugger` valeur de la `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` clé de Registre pour marquer votre complément pour le débogage.

Sinon, si vous avez utilisé un autre outil pour créer votre complément, effectuez les étapes suivantes.

1. Accédez à la clé de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` Registre. Remplacez `[Add-in ID]` par **l’ID** du manifeste de votre complément.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Définissez la valeur de la `UseDirectDebugger` clé sur `1`.

## <a name="configure-and-run-the-debugger"></a>Configurer et exécuter le débogueur

Maintenant que vous avez activé le débogage sur votre complément, vous êtes prêt à configurer et exécuter le débogueur. Pour obtenir des instructions sur la façon de procéder, sélectionnez l’une des options suivantes qui s’applique à votre runtime.

- Si votre complément s’exécute dans le runtime WebView, [reportez-vous à Microsoft Office extension de débogueur de complément pour Visual Studio Code](../testing/debug-with-vs-extension.md).

- Si votre complément s’exécute dans le runtime Microsoft Edge Chromium WebView2, [reportez-vous à Debug add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (basé sur Chromium).](../testing/debug-desktop-using-edge-chromium.md)

## <a name="see-also"></a>Voir aussi

- [Commandes de complément pour Outlook](add-in-commands-for-outlook.md)
- [Vue d’ensemble du débogage Office des modules](../testing/debug-add-ins-overview.md)
- [Déboguer votre complément Outlook basé sur des événements](debug-autolaunch.md)
