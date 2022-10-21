---
title: Déboguer votre complément Outlook basé sur les événements
description: Découvrez comment déboguer votre complément Outlook qui implémente l’activation basée sur les événements.
ms.topic: article
ms.date: 10/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: e8065c454bbe1587a6e5b7189a4522c229e9aed1
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664671"
---
# <a name="debug-your-event-based-outlook-add-in"></a>Déboguer votre complément Outlook basé sur les événements

Cet article fournit des conseils de débogage lorsque vous implémentez [l’activation basée sur les événements](autolaunch.md) dans votre complément. La fonctionnalité d’activation basée sur les événements a été introduite dans [l’ensemble de conditions requises 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), avec des événements supplémentaires désormais disponibles dans les ensembles de conditions requises suivants. Pour plus d’informations, consultez [Événements pris en charge](autolaunch.md#supported-events).

> [!IMPORTANT]
> Cette fonctionnalité de débogage est uniquement prise en charge dans Outlook sur Windows avec un abonnement Microsoft 365.

Cet article décrit les étapes clés pour activer le débogage.

- [Marquer le complément pour le débogage](#mark-your-add-in-for-debugging)
- [Configurer Visual Studio Code](#configure-visual-studio-code)
- [Attacher Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

Si vous avez utilisé le générateur Yeoman pour les compléments Office pour créer votre projet de complément (par exemple, en effectuant la [procédure pas à pas d’activation basée sur les événements](autolaunch.md)), suivez l’option **Créé avec le générateur Yeoman** tout au long de cet article. Sinon, suivez les **étapes Autres** . Visual Studio Code doit être au moins version 1.56.1.

## <a name="mark-your-add-in-for-debugging"></a>Marquer votre complément pour le débogage

1. Définissez la clé `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`de Registre . Remplacez par `[Add-in ID]` l’ID du complément à partir du manifeste.

    - **Manifeste XML** : utilisez la valeur de l’élément **\<Id\>** enfant de l’élément racine **\<OfficeApp\>** .
    - **Manifeste Teams (préversion)** : utilisez la valeur de la propriété « id » de l’objet anonyme `{ ... }` racine.

    **Créé avec le générateur Yeoman** : dans une fenêtre de ligne de commande, accédez à la racine de votre dossier de complément, puis exécutez la commande suivante.

    ```command&nbsp;line
    npm start
    ```

    En plus de générer le code et de démarrer le serveur local, cette commande doit définir la clé de `UseDirectDebugger` Registre pour ce complément sur `1`.

    **Autre** : ajoutez la clé de `UseDirectDebugger` Registre sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`. Remplacez par `[Add-in ID]` le **\<Id\>** à partir du manifeste du complément. Définissez la clé de Registre sur `1`.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Démarrez Outlook ou redémarrez-le s’il est déjà ouvert.
1. Composez un nouveau message ou un nouveau rendez-vous. La boîte de dialogue Debug Event-based handler (Debug Event-based handler) doit s’afficher. *N’interagissez pas* encore avec la boîte de dialogue.

    ![Boîte de dialogue Debug Event-based handler dans Windows.](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Configurer Visual Studio Code

### <a name="created-with-yeoman-generator"></a>Créé avec le générateur Yeoman

1. De retour dans la fenêtre de ligne de commande, ouvrez Visual Studio Code.

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
1. Accédez à **Fichier** > **Ouvrir le dossier**, accédez au dossier que vous venez de créer, puis choisissez **Sélectionner un dossier**.
1. Dans la barre d’activité, sélectionnez **Exécuter et déboguer** (Ctrl+Maj+D).

    ![Icône Exécuter et déboguer dans la barre d’activité.](../images/vs-code-debug.png)

1. Sélectionnez le lien **Créer un fichier launch.json** .

    ![Lien situé sous l’option Exécuter et déboguer pour créer un fichier launch.json dans Visual Studio Code.](../images/vs-code-create-launch.json.png)

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

1. Pour rechercher le **bundle.js** du complément , ouvrez le dossier suivant dans l’Explorateur Windows et recherchez les fichiers de **\<Id\>** votre complément (trouvés dans le manifeste).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Ouvrez le dossier précédé de cet ID et copiez son chemin d’accès complet. Dans Visual Studio Code, ouvrez **bundle.js** à partir de ce dossier. Le modèle du chemin d’accès au fichier doit être le suivant :

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Placez les points d’arrêt dans bundle.js où vous souhaitez que le débogueur s’arrête.
1. Dans la liste déroulante **DEBUG** , sélectionnez **Débogage direct**, puis sélectionnez l’icône **Démarrer le débogage** .

    ![Option Débogage direct sélectionnée dans les options de configuration dans la liste déroulante Débogage de Visual Studio Code.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Débogage

1. Après avoir vérifié que le débogueur est attaché, revenez à Outlook et, dans la boîte de dialogue **Déboguer le gestionnaire basé sur les** événements, choisissez **OK** .

1. Vous pouvez maintenant atteindre vos points d’arrêt dans Visual Studio Code, ce qui vous permet de déboguer votre code d’activation basé sur les événements.

## <a name="stop-debugging"></a>Arrêter le débogage

Pour arrêter le débogage pour le reste de la session de bureau Outlook actuelle, dans la boîte de dialogue **Déboguer le gestionnaire basé sur les** événements, choisissez **Annuler**. Pour réactiver le débogage, redémarrez le bureau Outlook.

Pour empêcher la boîte de dialogue **Debug Event-based handler** de apparaître et d’arrêter le débogage pour les sessions Outlook suivantes, supprimez la clé de Registre associée ou définissez sa valeur `0`sur : `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.

## <a name="see-also"></a>Voir aussi

- [Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md)
- [Déboguer votre complément avec la journalisation runtime](../testing/runtime-logging.md#runtime-logging-on-windows)
