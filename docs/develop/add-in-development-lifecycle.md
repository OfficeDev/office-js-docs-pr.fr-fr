
# <a name="office-add-ins-development-lifecycle"></a>Cycle de vie du développement des compléments Office

>
  **Remarque :** Lorsque vous créez votre complément, si vous envisagez de le [publier](../publish/publish.md) dans Office Store, assurez-vous que vous respectez les [stratégies de validation Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability)).

Le cycle de vie de développement classique d’un complément Office comprend les étapes suivantes :


1.  **Déterminez l’objet du complément.**
    
    Posez-vous les questions suivantes :
    
      - Quelle peut être l’utilité du complément ? 
    
      - Comment peut-elle contribuer à accroître la productivité de vos clients ?
    
      - Quels scénarios sont pris en charge par les fonctionnalités de votre complément ?
    

    Déterminez les fonctionnalités et les scénarios les plus importants et réalisez la conception à partir de ces éléments. 
    
2.  **Identifiez les données et la source de données du complément.**
    
    Les données proviennent-elles d’un document, d’un classeur, d’une présentation, d’un projet ou sont-elles accessibles via un navigateur Access, ou bien proviennent-elles d’un ou de plusieurs éléments d’une boîte aux lettres Exchange Server ou Exchange Online ? Les données proviennent-elles d’une source externe telle qu’un service web ?
    
3.  **Identifiez le type de complément et les applications hôtes Office les mieux adaptés pour prendre en charge l’objet de l’application.**
    
    Tenez compte des informations suivantes pour identifier les scénarios :
    
    - Les clients utiliseront-ils le complément pour enrichir le contenu d’un document ou d’une base de données Access basée sur navigateur ? Si c’est le cas, vous pouvez envisager de créer un complément de contenu. 
    
    - Les clients utiliseront-ils le complément lors de la visualisation ou de la composition d’un message électronique ou d’un rendez-vous ? Est-il important de pouvoir exposer le complément conformément au contexte actuel ? La possibilité de rendre le complément disponible non seulement sur le bureau mais également sur des tablettes ou des smartphones constitue-telle une priorité ?
    
        Si vous répondez par Oui à une de ces questions, envisagez de créer un complément Outlook. Ensuite, identifiez le contexte qui déclenchera votre complément (par exemple, un formulaire de composition utilisé par un utilisateur, des types de messages spécifiques, la présence d’une pièce jointe, l’adresse, la suggestion de tâche, la suggestion de réunion ou certains modèles de chaînes dans le contenu d’un courrier électronique ou d’un rendez-vous). Reportez-vous à l’article relatif aux [règles d’Activation pour les compléments Outlook](../outlook/manifests/activation-rules.md) pour savoir comment activer le complément Outlook en fonction du contexte.
    
    - Les clients utiliseront-ils le complément pour améliorer l’affichage ou l’expérience de création d’un document ? Si c’est le cas, vous pouvez créer un complément de volet Office. 

    La prise en charge pour certaines API de complément peut-être différente entre les applications Office et la plateforme d’exécution (Windows, Mac, Web, Mobile). Pour afficher la couverture API actuelle par le client et la plateforme, consultez la page concernant la [disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability).  
    
4.  **Concevez et implémentez l’expérience utilisateur et l’interface utilisateur pour le complément.**
    
    Concevez une expérience utilisateur rapide et fluide qui soit cohérente, facile à apprendre, avec des scénarios nécessitant uniquement quelques étapes d’exécution. Selon l’objet du complément, utilisez des API ou des services web de tiers.
    
    Vous pouvez faire votre choix parmi divers outils de développement web et utiliser du code HTML et JavaScript pour implémenter l’interface utilisateur.
    
5.  **Créez un fichier manifeste XML basé sur le schéma de manifeste des compléments Office.**
    
    Créez un manifeste XML pour identifier le complément et sa configuration requise, spécifiez l’emplacement des fichiers HTML, JavaScript et CSS que le complément utilise et, selon le type du complément, la taille et les autorisations par défaut.
    
    Pour les compléments Outlook, vous pouvez spécifier le contexte, en fonction du message ou du rendez-vous actif, sous lequel votre complément est pertinent et doit être disponible dans l’interface utilisateur d’Outlook. Vous devez également choisir les périphériques que votre complément doit prendre en charge. Dans le manifeste, spécifiez le contexte sous forme de règles d’activation, ainsi que les périphériques pris en charge.
    
6.  **Installez et testez le complément.**
    
    Placez les fichiers HTML et les éventuels fichiers JavaScript et CSS sur les serveurs web qui sont spécifiés dans le fichier manifeste du complément. Le processus d’installation d’un complément dépend du type de celui-ci. Pour plus d’informations, reportez-vous à la page relative au [chargement d’une version test des compléments Office à des fins de test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
    
    Pour les compléments Outlook, installez-le dans une boîte aux lettres Exchange et spécifiez l’emplacement du fichier manifeste du complément dans le Centre d’administration Exchange (CAE). Pour plus d’informations, consultez la rubrique [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md).
    
7.  **Publiez le complément.**
    
    Vous pouvez envoyer le complément à l’Office Store, à partir duquel les clients peuvent installer le complément. En outre, vous pouvez publier le volet des tâches et les compléments du contenu dans le catalogue de compléments d’un dossier privé dans SharePoint ou dans un dossier réseau partagé, et vous pouvez déployer un complément Outlook directement sur un serveur Exchange pour votre organisation. Pour plus d’informations, consultez la rubrique [Publier votre complément Office](../publish/publish.md).
    
8.  **Mettez à jour le complément.**
    
    Si votre complément lance un service web, et si vous apportez des mises à jour au service web après avoir publié le complément, vous n’avez pas besoin de republier le complément. Toutefois, si vous modifiez des éléments ou des données que vous avez soumis pour votre complément, comme le fichier manifeste du complément, les captures d’écran, les icônes, les fichiers HTML ou JavaScript, vous devrez republier le complément. Si vous avez notamment publié le complément dans l’Office Store, vous devrez soumettre à nouveau votre complément pour qu’Office Store puisse appliquer ces modifications. Vous devez renvoyer votre complément avec une mise à jour du fichier manifeste du complément qui inclut un nouveau numéro de version. Vous devez également veiller à mettre à jour le numéro de version du complément dans le formulaire d’envoi pour qu’il corresponde au numéro de version du nouveau fichier manifeste. Pour les compléments Outlook, vous devez vous assurer que l’élément [Id](../../reference/manifest/id.md) contient un UUID différent dans le fichier manifeste du complément.
    
