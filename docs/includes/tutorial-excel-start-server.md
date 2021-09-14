Si le serveur web local est déjà en cours d’exécution et que votre complément est déjà chargé dans Excel, passez à l’étape 2. Sinon, démarrez le serveur web local et chargez la version test de votre complément : 

- Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet. Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Excel avec votre complément chargé.

    ```command&nbsp;line
    npm start
    ```

- Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).

    ```command&nbsp;line
    npm run start:web
    ```

    Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).
