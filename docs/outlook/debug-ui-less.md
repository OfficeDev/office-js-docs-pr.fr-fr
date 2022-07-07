---
title: Déboguer votre complément Outlook sans interface utilisateur
description: Découvrez comment déboguer votre complément Outlook sans interface utilisateur.
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: e46bdf15172f5224995b17c39df4ba60ca6380ad
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660206"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>Déboguer votre complément Outlook sans interface utilisateur

Cet article explique comment utiliser l’extension de débogueur de complément Office dans Visual Studio Code pour déboguer [des compléments Outlook sans interface utilisateur](add-in-commands-for-outlook.md#executing-a-javascript-function). Les actions de complément sans interface utilisateur sont lancées via un bouton de commande de complément dans le ruban. Pour plus d’informations sur les commandes de complément, consultez [commandes de complément pour Outlook](add-in-commands-for-outlook.md).

Cet article part du principe que vous disposez déjà d’un projet de complément que vous souhaitez déboguer. Pour créer un complément sans interface utilisateur pour effectuer le débogage, suivez les étapes du [didacticiel : Créer un complément Outlook de composition de message](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Marquer votre complément pour le débogage

Si vous avez utilisé le [générateur Yeoman pour les compléments Office pour](../develop/yeoman-generator-overview.md) créer votre projet de complément, passez à la section [Configurer et exécutez le débogueur](#configure-and-run-the-debugger) plus loin dans cet article. Lorsque vous exécutez `npm start` pour générer votre complément et démarrer le serveur local, la commande définit également la `UseDirectDebugger` valeur de la `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` clé de Registre pour marquer votre complément pour le débogage.

Sinon, si vous avez utilisé un autre outil pour créer votre complément, effectuez les étapes suivantes.

1. Accédez à la clé de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` Registre. Remplacez `[Add-in ID]` par le **\<Id\>** manifeste de votre complément.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Définissez la valeur de la `UseDirectDebugger` clé sur `1`.

## <a name="configure-and-run-the-debugger"></a>Configurer et exécuter le débogueur

Maintenant que vous avez activé le débogage sur votre complément, vous êtes prêt à configurer et exécuter le débogueur. Pour obtenir des instructions sur la façon de procéder, sélectionnez l’une des options suivantes qui s’applique à votre runtime.

- Si votre complément s’exécute dans le runtime WebView, [reportez-vous à l’extension de débogueur de complément Microsoft Office pour Visual Studio Code](../testing/debug-with-vs-extension.md).

- Si votre complément s’exécute dans le runtime Microsoft Edge Chromium WebView2, [reportez-vous aux compléments de débogage sur Windows à l’aide de Visual Studio Code et de Microsoft Edge WebView2 (basé sur Chromium).](../testing/debug-desktop-using-edge-chromium.md)

## <a name="see-also"></a>Voir aussi

- [Commandes de complément pour Outlook](add-in-commands-for-outlook.md)
- [Vue d’ensemble du débogage Office des modules](../testing/debug-add-ins-overview.md)
- [Déboguer votre complément Outlook basé sur les événements](debug-autolaunch.md)
