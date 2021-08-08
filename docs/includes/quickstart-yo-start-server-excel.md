
Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.

> [!NOTE]
> Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.

> [!TIP]
> Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer. Lorsque vous exécutez cette commande, le serveur web local démarre.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- Installez les dépendances de votre complément dans le répertoire racine de votre projet.

     ```command&nbsp;line
    npm install
    ```

- Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet. Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Excel avec votre complément chargé.

    ```command&nbsp;line
    npm start
    ```

- Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

    ```command&nbsp;line
    npm run start:web
    ```

    Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de la version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

