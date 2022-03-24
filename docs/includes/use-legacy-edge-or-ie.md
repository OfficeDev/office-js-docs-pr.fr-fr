Si votre projet est basé sur node.js (c’est-à-dire, non développé avec Visual Studio et le serveur d’informations Internet (IIS), vous pouvez forcer Office sur Windows à utiliser edge hérité ou Internet Explorer pour exécuter des modules complémentaires, même si vous avez une combinaison de versions de Windows et de Office qui utiliserait normalement un navigateur plus récent. Pour plus d’informations sur les navigateurs utilisés par différentes combinaisons de versions Windows et Office, voir [Navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les Office de recherche.

1. Si votre projet n’a pas été créé avec le générateur [Yeoman pour l’outil Office Add-ins](../develop/yeoman-generator-overview.md), vous devez installer l’outil office-addin-dev-settings. Exécutez la commande suivante dans une invite de commandes.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Spécifiez le navigateur que vous Office utiliser avec la commande suivante dans une invite de commandes à la racine du projet. Remplacez `<path-to-manifest>` par le chemin d’accès relatif, qui est simplement le nom du fichier manifeste s’il se trouve à la racine du projet. Remplacez `<webview>` par l’un ou `ie` l’autre .`edge-legacy`

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    Voici un exemple.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    Vous devriez voir un message dans la ligne de commande que le type de vue web est désormais définie sur IE (ou edge hérité).

1. Lorsque vous avez terminé, définissez Office pour reprendre l’utilisation du navigateur par défaut pour votre combinaison de versions Windows et Office à l’aide de la commande suivante.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
