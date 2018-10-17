1. Ouvrez un terminal bash à la racine du projet (**[...] /Mes complément office**) et exécutez la commande suivante pour démarrer le serveur de développement.

    ```bash
    npm start
    ```

2. Ouvrez Internet Explorer ou Microsoft Edge et accédez à la page `https://localhost:3000`. Si la page se charge sans erreurs de certificat, passez à la section suivante de cet article (**Essayez**). Si votre navigateur indique que le certificat de site n’est pas approuvé, passez à l’étape suivante.

3. Les compléments web Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez. Si votre navigateur indique que le certificat de site n’est pas approuvé, vous devez ajouter le certificat en tant que certificat approuvé. Consultez l’article relatif à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > Il est possible que le navigateur web Chrome continue d’indiquer que le certificat du site n’est pas approuvé, même si vous avez suivi les étapes décrites dans l’article relatif à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Par conséquent, vous devez utiliser Internet Explorer ou Microsoft Edge pour vérifier que le certificat est approuvé. 

4. Une fois que votre navigateur a chargé la page du complément sans erreurs de certificat, vous pouvez tester votre complément.
