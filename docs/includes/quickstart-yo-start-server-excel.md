
Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.

> [!NOTE]
> Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.

> [!TIP]
> Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer. Lorsque vous exécutez cette commande, le serveur web local démarre.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas encore en cours d’exécution), et Excel s’ouvre avec votre complément chargé.

    ```command&nbsp;line
    npm start
    ```

- Pour tester votre complément dans Excel Online, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

    ```command&nbsp;line
    npm run start:web
    ```

    Pour utiliser votre complément, ouvrez un nouveau document dans Excel Online, puis chargez indépendamment votre complément en suivant les instructions fournies dans [Chargement de version test d’un complément Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).

