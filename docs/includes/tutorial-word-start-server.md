Si le serveur web local est déjà en cours d’exécution et que votre complément est déjà chargé dans Word, passez à l’étape 2. Sinon, démarrez le serveur web local et chargez la version test de votre complément : 

- Pour tester votre complément dans Word, exécutez la commande suivante dans le répertoire racine de votre projet. Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Word avec votre complément chargé.

    ```command&nbsp;line
    npm start
    ```

- Pour tester votre complément dans Word sur le web, exécutez la commande suivante dans le répertoire racine de votre projet. Lorsque vous exécutez cette commande, le serveur web local démarre. Remplacez « {url} » par l’URL d’un document Word sur votre OneDrive ou une bibliothèque SharePoint sur laquelle vous avez des autorisations.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

