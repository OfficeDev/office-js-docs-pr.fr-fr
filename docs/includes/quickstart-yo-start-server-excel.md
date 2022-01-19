
Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.

[!INCLUDE [alert use https](alert-use-https.md)]

> [!TIP]
> Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer. Lorsque vous exécutez cette commande, le serveur web local démarre.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet. Cela démarre le serveur web local et ouvre Excel avec votre complément chargé.

    ```command&nbsp;line
    npm start
    ```

- Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre. Remplacez « {url} » par l’URL d’un document Excel sur votre OneDrive ou une bibliothèque SharePoint sur laquelle vous disposez d’autorisations.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]
