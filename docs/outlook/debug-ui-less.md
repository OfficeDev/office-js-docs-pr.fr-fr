---
title: Commandes de fonction de débogage dans les compléments Outlook
description: Découvrez comment déboguer des commandes de fonction dans des compléments Outlook.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6189824fd526d48321b355c9b306fa5ef732f411
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797588"
---
# <a name="debug-function-commands-in-outlook-add-ins"></a>Commandes de fonction de débogage dans les compléments Outlook

> [!NOTE]
> La technique décrite dans cet article ne peut être utilisée que sur un ordinateur de développement Windows. Si vous développez sur un Mac, consultez [les commandes de fonction de débogage](../testing/debug-function-command.md).

Cet article explique comment utiliser l’extension de débogueur de complément Office dans Visual Studio Code pour déboguer [des commandes de fonction](add-in-commands-for-outlook.md#run-a-function-command). Les commandes de fonction sont lancées via un bouton de commande de complément dans le ruban. Pour plus d’informations sur les commandes de complément, consultez [commandes de complément pour Outlook](add-in-commands-for-outlook.md).

Cet article part du principe que vous disposez déjà d’un projet de complément que vous souhaitez déboguer. Pour créer un complément avec une commande de fonction pour effectuer le débogage, suivez les étapes du [didacticiel : créer un complément Outlook de composition de message](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Marquer votre complément pour le débogage

Si vous avez utilisé le [générateur Yeoman pour les compléments Office pour](../develop/yeoman-generator-overview.md) créer votre projet de complément, passez à la section [Configurer et exécutez le débogueur](#configure-and-run-the-debugger) plus loin dans cet article. Lorsque vous exécutez `npm start` pour générer votre complément et démarrer le serveur local, la commande définit également la `UseDirectDebugger` valeur de la `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` clé de Registre pour marquer votre complément pour le débogage.

Sinon, si vous avez utilisé un autre outil pour créer votre complément, effectuez les étapes suivantes.

1. Accédez à la clé de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` Registre. Remplacez `[Add-in ID]` par le **\<Id\>** manifeste de votre complément.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Définissez la valeur de la `UseDirectDebugger` clé sur `1`.

## <a name="configure-and-run-the-debugger"></a>Configurer et exécuter le débogueur

Maintenant que vous avez activé le débogage sur votre complément, vous êtes prêt à configurer et exécuter le débogueur. Pour obtenir des instructions sur la façon de procéder, sélectionnez l’une des options suivantes qui s’applique à votre contrôle webview. Pour plus d’informations sur la façon de déterminer le contrôle webview utilisé sur votre ordinateur de développement, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

- Si votre complément s’exécute dans le contrôle webview incorporé à partir de Edge Legacy (EdgeHTML), consultez [l’extension de débogueur de complément Microsoft Office pour Visual Studio Code](../testing/debug-with-vs-extension.md).

- Si votre complément s’exécute dans le contrôle webview incorporé à partir de Microsoft Edge Chromium (WebView2), consultez [Les compléments de débogage sur Windows à l’aide de Visual Studio Code et de Microsoft Edge WebView2 (basé sur Chromium).](../testing/debug-desktop-using-edge-chromium.md)

## <a name="see-also"></a>Voir aussi

- [Commandes de complément pour Outlook](add-in-commands-for-outlook.md)
- [Vue d’ensemble du débogage Office des modules](../testing/debug-add-ins-overview.md)
- [Déboguer votre complément Outlook basé sur les événements](debug-autolaunch.md)
