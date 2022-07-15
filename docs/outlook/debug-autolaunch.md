---
title: Déboguer votre complément Outlook basé sur les événements
description: Découvrez comment déboguer votre complément Outlook qui implémente l’activation basée sur les événements.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5d36a23b34132071077e3eb192e562288befb8a5
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797490"
---
# <a name="debug-your-event-based-outlook-add-in"></a>Déboguer votre complément Outlook basé sur les événements

Cet article fournit des conseils de débogage lorsque vous implémentez [l’activation basée sur les événements](autolaunch.md) dans votre complément. La fonctionnalité d’activation basée sur les événements a été introduite dans [l’ensemble de conditions requises 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10) avec des événements supplémentaires désormais disponibles en préversion. Pour plus d’informations, [reportez-vous aux événements pris en charge](autolaunch.md#supported-events).

> [!IMPORTANT]
> Cette fonctionnalité de débogage est uniquement prise en charge dans Outlook sur Windows avec un abonnement Microsoft 365.

Dans cet article, nous abordons les étapes clés pour activer le débogage.

- [Marquer le complément pour le débogage](#mark-your-add-in-for-debugging)
- [Configurer Visual Studio Code](#configure-visual-studio-code)
- [Attacher Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

Si vous avez utilisé yeoman Generator pour compléments Office pour créer votre projet de complément (par exemple, en effectuant la [procédure pas à pas d’activation basée sur les événements](autolaunch.md)), suivez l’option **Créer avec le générateur Yeoman** tout au long de cet article. Sinon, suivez les **autres** étapes. Visual Studio Code doit être au moins version 1.56.1.

## <a name="mark-your-add-in-for-debugging"></a>Marquer votre complément pour le débogage

1. Définissez la clé `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`de Registre . `[Add-in ID]` est le **\<Id\>** manifeste du complément.

    **Créé avec le générateur Yeoman** : dans une fenêtre de ligne de commande, accédez à la racine de votre dossier de complément, puis exécutez la commande suivante.

    ```command&nbsp;line
    npm start
    ```

    En plus de générer le code et de démarrer le serveur local, cette commande doit définir la `UseDirectDebugger` clé de Registre pour ce complément `1`sur .

    **Autre** : ajoutez la clé de `UseDirectDebugger` Registre sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`. Remplacez `[Add-in ID]` par le **\<Id\>** manifeste du complément. Définissez la clé `1`de Registre sur .

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Démarrez le bureau Outlook (ou redémarrez Outlook s’il est déjà ouvert).
1. Composez un nouveau message ou rendez-vous. La boîte de dialogue suivante doit s’afficher. *N’interagissez pas* encore avec la boîte de dialogue.

    ![Capture d’écran de la boîte de dialogue Gestionnaire basé sur les événements de débogage.](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Configurer Visual Studio Code

### <a name="created-with-yeoman-generator"></a>Créé avec le générateur Yeoman

1. Dans la fenêtre de ligne de commande, ouvrez Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. Dans Visual Studio Code, ouvrez le fichier **./.vscode/launch.json** et ajoutez l’extrait suivant à votre liste de configurations. Enregistrez vos modifications.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a>Autre

1. Créez un dossier appelé **Débogage** (peut-être dans votre dossier **Bureau** ).
1. Ouvrez Visual Studio Code.
1. Accédez au **dossier Ouvrir** un **fichier** > , accédez au dossier que vous venez de créer, puis **sélectionnez Sélectionner un dossier**.
1. Dans la barre d’activité, sélectionnez l’élément **de débogage** (Ctrl+Maj+D).

    ![Capture d’écran de l’icône Déboguer dans la barre d’activité.](../images/vs-code-debug.png)

1. Sélectionnez le lien **créer un fichier launch.json** .

    ![Capture d’écran du lien permettant de créer un fichier launch.json dans Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. Dans la liste déroulante **Sélectionner un environnement** , sélectionnez **Edge : Lancer** pour créer un fichier launch.json.
1. Ajoutez l’extrait suivant à votre liste de configurations. Enregistrez vos modifications.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a>Attacher Visual Studio Code

1. Pour trouver le **bundle.js** du complément, ouvrez le dossier suivant dans l’Explorateur Windows et recherchez celui de **\<Id\>** votre complément (trouvé dans le manifeste).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Ouvrez le dossier préfixé avec cet ID et copiez son chemin d’accès complet. Dans Visual Studio Code, ouvrez **bundle.js** à partir de ce dossier. Le modèle du chemin d’accès au fichier doit être le suivant :

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Placez les points d’arrêt dans bundle.js où vous souhaitez que le débogueur s’arrête.
1. Dans la liste **déroulante DEBUG** , sélectionnez le nom **Débogage direct**, puis sélectionnez **Exécuter**.

    ![Capture d’écran de la sélection du débogage direct dans les options de configuration dans la liste déroulante Débogage Visual Studio Code.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Débogage

1. Après avoir confirmé que le débogueur est attaché, revenez à Outlook et, dans la boîte **de dialogue Gestionnaire d’événements de débogage** , choisissez **OK** .

1. Vous pouvez maintenant atteindre vos points d’arrêt dans Visual Studio Code, ce qui vous permet de déboguer votre code d’activation basé sur les événements.

## <a name="stop-debugging"></a>Arrêter le débogage

Pour arrêter le débogage pour le reste de la session de bureau Outlook actuelle, dans la boîte de dialogue **Gestionnaire d’événements de débogage** , choisissez **Annuler**. Pour réactiver le débogage, redémarrez le bureau Outlook.

Pour empêcher que la boîte de dialogue gestionnaire **basé sur les événements de débogage** ne se déclenche et arrête le débogage pour les sessions Outlook suivantes, supprimez la clé de Registre associée ou définissez sa valeur `0`sur : `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md)
- [Déboguer votre complément avec la journalisation runtime](../testing/runtime-logging.md#runtime-logging-on-windows)
