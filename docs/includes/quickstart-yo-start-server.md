1. Ouvrez un terminal bash à la racine du projet (**[...] /Mes complément office**) et exécutez la commande suivante pour démarrer le serveur de développement.

    ```bash
    npm start
    ```

    Cela lance un serveur web sur `https://localhost:3000` et ouvre votre navigateur par défaut sur cette adresse.

2. Les compléments web Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si votre navigateur indique que le certificat de site n’est pas approuvé, vous devez ajouter le certificat en tant que certificat approuvé. Consultez l’article relatif à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > Il est possible que le navigateur web Chrome continue d’indiquer que le certificat du site n’est pas approuvé, même si vous avez suivi les étapes décrites dans l’article relatif à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Vous pouvez ignorer ce message d’avertissement dans Chrome. Vérifiez tout de même que le certificat est approuvé en entrant `https://localhost:3000` dans Internet Explorer ou Microsoft Edge. 

3. Une fois que votre navigateur a chargé la page du complément sans erreurs de certificat, vous pouvez tester votre complément. 
