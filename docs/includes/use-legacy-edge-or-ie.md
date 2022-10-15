Si votre projet est basé sur node.js (c’est-à-dire qu’il n’est pas développé avec Visual Studio et le serveur d’informations Internet (IIS),vous pouvez forcer Office sur Windows à utiliser Edge Legacy ou Internet Explorer pour exécuter des compléments, même si vous avez une combinaison de versions de Windows et d’Office qui utiliseraient normalement un navigateur plus récent. Pour plus d’informations sur les navigateurs utilisés par différentes combinaisons de versions De Windows et Office, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!NOTE]
> L’outil utilisé pour forcer la modification dans le navigateur est pris en charge uniquement dans le canal d’abonnement bêta de Microsoft 365. Rejoignez le [programme Office Insider](https://insider.office.com/join/windows) et sélectionnez l’option **Canal bêta** pour accéder aux versions Bêta d’Office. Voir aussi [À propos d’Office : Quelle version d’Office utilise-t-on ?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
>
> Strictement, c’est le `webview` commutateur de cet outil (voir **l’étape 2**) qui nécessite le canal bêta. L’outil a d’autres commutateurs qui n’ont pas cette exigence.

1. Si votre projet *n’a pas* été créé avec l’outil [générateur Yeoman pour compléments Office](../develop/yeoman-generator-overview.md) , vous devez installer l’outil office-addin-dev-settings. Exécutez la commande suivante dans une invite de commandes.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Spécifiez le navigateur que vous souhaitez qu’Office utilise avec la commande suivante dans une invite de commandes à la racine du projet. Remplacez `<path-to-manifest>` par le chemin d’accès relatif, qui est simplement le nom du fichier manifeste s’il se trouve à la racine du projet. Remplacez par `<webview>` l’un ou l’autre`ie`.`edge-legacy`

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    Voici un exemple.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    Vous devez voir un message dans la ligne de commande indiquant que le type de vue web est maintenant défini sur Internet Explorer (ou Edge Legacy).

1. Lorsque vous avez terminé, définissez Office pour qu’il reprenne à l’aide du navigateur par défaut pour votre combinaison de versions Windows et Office avec la commande suivante.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
