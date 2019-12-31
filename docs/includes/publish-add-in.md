Un complément Office comprend une application Web et un fichier manifeste. L’application Web définit l’interface utilisateur et les fonctionnalités du complément, tandis que le manifeste spécifie l’emplacement de l’application Web et définit les paramètres et fonctionnalités du complément. 

Lorsque vous développez votre complément, vous pouvez l’exécuter sur votre serveur Web local (`localhost`), mais lorsque vous êtes prêt à le publier pour permettre à d’autres utilisateurs d’y accéder, vous devez déployer l’application Web sur un serveur Web ou un service d’hébergement Web (par exemple, Microsoft Azure), puis mettre à jour le manifeste pour spécifier l’URL de l’application déployée. 

Lorsque votre complément fonctionne comme vous le souhaitez et que vous êtes prêt à le publier pour permettre à d’autres utilisateurs d’y accéder, procédez comme suit :

1. À partir de la ligne de commande, dans le répertoire racine de votre projet de complément, exécutez la commande suivante pour préparer tous les fichiers pour le déploiement en production : 

    ```command&nbsp;line
    npm run build
    ```

    Une fois la build terminée, le dossier **dist** dans le répertoire racine de votre projet de complément contient les fichiers que vous allez déployer dans les étapes suivantes.

2. Chargez le contenu du dossier **dist** sur le serveur Web qui héberge votre complément. Vous pouvez utiliser n’importe quel type de serveur Web ou de service d’hébergement Web pour héberger votre complément.

3. Dans VS Code, ouvrez le fichier manifeste du complément situé dans le répertoire racine du projet (`manifest.xml`). Remplacez toutes les occurrences de `https://localhost:3000` par l’URL de l’application Web que vous avez déployée sur un serveur Web à l’étape précédente.

4. Choisissez la méthode que vous voulez utiliser pour [déployer votre complément Office](../publish/publish.md), puis suivez les instructions pour publier le fichier manifeste.
