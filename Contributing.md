# <a name="contribute-to-this-documentation"></a>Contribuer à cette documentation

Nous vous remercions de l’intérêt que vous portez à notre documentation.

* [Méthodes de collaboration](#ways-to-contribute)
* [Contribuer à l’aide de GitHub](#contribute-using-github)
* [Contribuer à l’aide de Git](#contribute-using-git)
* [Utilisation de Markdown pour mettre en forme votre rubrique](#how-to-use-markdown-to-format-your-topic)
* [FAQ](#faq)
* [Autres ressources](#more-resources)

## <a name="ways-to-contribute"></a>Méthodes de collaboration

Voici comment vous pouvez contribuer à cette documentation :

* Pour apporter des modifications mineures à un article, [contribuez à l’aide de GitHub](#contribute-using-github).
* Pour apporter des modifications importantes, ou des modifications impliquant du code, [contribuez à l’aide de git](#contribute-using-git).
* Signalez les bogues de documentation via des problèmes GitHub.
* Demandez une nouvelle documentation sur le site [Office Developer Platform UserVoice](http://officespdev.uservoice.com) .

## <a name="contribute-using-github"></a>Contribuer à l’aide de GitHub

Utilisez GitHub pour contribuer à cette documentation sans avoir à cloner le référentiel sur votre ordinateur de bureau. Il s’agit de la méthode la plus simple pour créer une requête de tirage dans ce référentiel. Utilisez cette méthode pour apporter une modification mineure qui n’implique pas de modifications de code. 

**Remarque**: cette méthode vous permet de contribuer à un article à la fois.

### <a name="to-contribute-using-github"></a>Pour contribuer à l’aide de GitHub

1. Recherchez l’article auquel vous souhaitez contribuer sur GitHub.
2. Une fois que vous êtes sur l’article dans GitHub, connectez-vous à GitHub (obtenir un compte gratuit [rejoindre GitHub](https://github.com/join)).
3. Choisissez l' **icône** en forme de crayon (modifiez le fichier dans votre bifurcation de ce projet) et effectuez vos modifications dans la fenêtre **<>modifier le fichier** . 
4. Faites défiler jusqu’en bas et entrez une description.
5. Sélectionnez **proposer une modification de fichier** > **créer une requête de tirage**.

Vous avez maintenant envoyé une demande de tirage. Les requêtes de tirage sont généralement revues dans les 10 jours ouvrés. 


## <a name="contribute-using-git"></a>Contribuer à l’aide de Git

Utilisez git pour apporter des modifications de fond, telles que :

* Code contributeur.
* Apporter des modifications qui ont une incidence sur la signification.
* Contribution aux modifications importantes apportées au texte.
* Ajout de nouvelles rubriques.

### <a name="to-contribute-using-git"></a>Pour contribuer à l’aide de git

1. Si vous n’avez pas de compte GitHub, configurez-en un sur [GitHub](https://github.com/join). 
2. Une fois que vous disposez d’un compte, installez git sur votre ordinateur. Suivez les étapes décrites dans le didacticiel de [configuration de git] .
3. Pour envoyer une requête de tirage à l’aide de git, suivez les étapes de la procédure d' [utilisation de GitHub, de git et de ce référentiel](#use-github-git-and-this-repository).
4. Vous serez invité à signer le contrat de licence du collaborateur si vous procédez comme suit :

    * Un membre du groupe Microsoft Open technologies.
    * Un collaborateur qui ne travaille pas pour Microsoft.

En tant que membre de la Communauté, vous devez signer le contrat de licence de contribution (CLA) avant de pouvoir contribuer à un projet. Vous n’avez besoin d’effectuer et de soumettre la documentation qu’une seule fois. Lisez attentivement le document. Il faudra peut-être que votre employeur signe le document.

La signature du CLA ne vous octroie pas de droits de validation sur le référentiel principal, mais cela signifie que les équipes de publication de contenu Office Developer et Office peuvent consulter et approuver vos contributions. Vous avez crédité pour vos envois.

Les requêtes de tirage sont généralement revues dans les 10 jours ouvrés.

## <a name="use-github-git-and-this-repository"></a>Utilisation de GitHub, de Git et du référentiel

**Remarque**: la plupart des informations contenues dans cette section sont disponibles dans les articles [d’aide GitHub] .  Si vous êtes familiarisé avec git et GitHub, passez à la section **contribution et modification du contenu** pour les caractéristiques du flux de code/contenu de ce référentiel.

### <a name="to-set-up-your-fork-of-the-repository"></a>Pour configurer votre bifurcation du référentiel

1.  Configurez un compte GitHub pour pouvoir contribuer à ce projet. Si vous ne l’avez pas fait, accédez à [GitHub](https://github.com/join) et faites-le maintenant.
2.  Installez Git sur votre ordinateur. Suivez les étapes décrites dans le didacticiel de [configuration de git] .
3.  Créez votre propre bifurcation du référentiel. Pour ce faire, en haut de la page, cliquez sur le bouton **bifurcation** .
4.  Copiez votre bifurcation sur votre ordinateur. Pour ce faire, ouvrez git bash. À l’invite de commandes, entrez les informations suivantes :

        git clone https://github.com/<your user name>/<repo name>.git

    Créez ensuite une référence vers le référentiel racine en entrant les commandes suivantes :

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Félicitations ! Votre référentiel est maintenant configuré. Vous n’aurez plus besoin d’effectuer ces étapes.

### <a name="contribute-and-edit-content"></a>Contribution et modification de contenu

Pour que le processus de contribution soit aussi transparent que possible, suivez les étapes ci-dessous.

#### <a name="to-contribute-and-edit-content"></a>Pour contribuer et modifier le contenu

1. Créez une branche.
2. Ajoutez du nouveau contenu ou modifiez le contenu existant.
3. Envoyez une requête de tirage au référentiel principal.
4. Supprimez la branche.

**Important** Limitez chaque succursale à un seul concept/article pour rationaliser le flux de travail et réduire le risque de conflits de fusion. Le contenu approprié pour une nouvelle succursale inclut :

* Un nouvel article.
* Modifications de l’orthographe et de la grammaire.
* Application d’un seul changement de mise en forme dans un grand ensemble d’articles (par exemple, application d’un nouveau pied de page de copyright).

#### <a name="to-create-a-new-branch"></a>Pour créer une nouvelle branche

1.  Ouvrez git bash.
2.  À l’invite de commandes git bash, tapez `git pull upstream master:<new branch name>` . Cela crée une nouvelle branche locale qui est copiée à partir de la branche principale OfficeDev la plus récente.
3.  À l’invite de commandes git bash, tapez `git push origin <new branch name>` . Cette alerte GitHub à la nouvelle branche. La nouvelle branche doit maintenant apparaître dans votre bifurcation du référentiel de GitHub.
4.  À l’invite de commandes git bash, saisissez `git checkout <new branch name>` pour basculer vers votre nouvelle branche.

#### <a name="add-new-content-or-edit-existing-content"></a>Ajout de nouveau contenu ou modification de contenu existant

Vous accédez au référentiel sur votre ordinateur à l’aide de l’Explorateur de fichiers. Les fichiers du référentiel se trouvent dans `C:\Users\<yourusername>\<repo name>` .

Pour modifier des fichiers, ouvrez-les dans un éditeur de votre choix et modifiez-les. Pour créer un fichier, utilisez l’éditeur de votre choix et enregistrez le nouveau fichier à l’emplacement approprié dans votre copie locale du référentiel. Tout en travaillant, enregistrez votre travail fréquemment.

Les fichiers dans `C:\Users\<yourusername>\<repo name>` constituent une copie de travail de la nouvelle branche que vous avez créée dans votre référentiel local. Si vous effectuez des modifications dans ce dossier, le référentiel local ne sera pas altéré tant que vous ne validez pas les modifications. Pour valider une modification dans le référentiel local, entrez les commandes suivantes dans GitBash :

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

La commande `add` ajoute vos modifications dans une zone de transit pour qu’elles soient validées dans le référentiel. La période qui suit la `add` commande spécifie que vous souhaitez préparer tous les fichiers que vous avez ajoutés ou modifiés, en vérifiant les sous-dossiers de manière récursive. (Si vous ne souhaitez pas valider toutes les modifications, vous pouvez ajouter des fichiers spécifiques. Vous pouvez également annuler une validation. Pour obtenir de l’aide, entrez `git add -help` ou `git status`.)

La commande `commit` applique les modifications en attente dans le référentiel. Le commutateur `-m` signifie que vous fournissez le commentaire de validation dans la ligne de commande. Les commutateurs-v et-a peuvent être omis. Le commutateur-v est destiné à une sortie détaillée de la commande, et-a fait ce que vous avez déjà fait avec la commande Add.

Vous pouvez valider plusieurs fois lorsque vous travaillez, ou vous pouvez effectuer une validation une fois que vous avez fini.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Envoyer une requête de tirage au référentiel principal

Lorsque vous avez terminé votre travail et que vous êtes prêt à le fusionner dans le référentiel principal, procédez comme suit.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Pour envoyer une requête de tirage au référentiel principal

1.  Dans l’invite de commandes git bash, tapez `git push origin <new branch name>` . Dans le référentiel local, `origin` fait référence au référentiel GitHub à partir duquel vous avez cloné le référentiel local. Cette commande affiche l’état actuel de votre nouvelle branche, y compris toutes les validations effectuées au cours des étapes précédentes dans votre bifurcation GitHub.
2.  Sur le site GitHub, accédez à la nouvelle branche dans votre bifurcation.
3.  Cliquez sur le bouton **requête de tirage** en haut de la page.
4.  Vérifiez que la branche de base est `OfficeDev/<repo name>@master` et que la branche de tête est `<your username>/<repo name>@<branch name>` .
5.  Sélectionnez le bouton **mettre à jour la plage de validation** .
6.  Ajoutez un titre à votre requête de tirage et décrivez toutes les modifications que vous effectuez.
7.  Envoyez la requête de tirage.

Un des administrateurs de site traitera votre demande de tirage. Votre requête de tirage doit apparaître sur le OfficeDev/ <repo name> site sous des problèmes. Lorsque la requête de tirage est acceptée, le problème est résolu.

#### <a name="create-a-new-branch-after-merge"></a>Créer une branche après fusion

Une fois la branche correctement fusionnée (autrement dit, votre demande de tirage est acceptée), ne continuez pas à travailler dans cette branche locale. Cela peut entraîner des conflits de fusion si vous soumettez une autre demande de tirage. Pour effectuer une autre mise à jour, créez une nouvelle branche locale à partir de la branche amont fusionnée, puis supprimez votre branche locale initiale.

Par exemple, si votre branche locale X a été correctement fusionnée avec la branche principale OfficeDev/Microsoft-Graph-docs et que vous souhaitez effectuer d’autres mises à jour du contenu qui a été fusionné. Créez une nouvelle branche locale, x2, à partir de la branche principale OfficeDev/Microsoft-Graph-docs. Pour ce faire, ouvrez GitBash et exécutez les commandes suivantes :

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Vous disposez maintenant de copies locales (dans une nouvelle branche locale) du travail que vous avez soumis dans la branche X. La branche x2 contient également tout le travail que d’autres rédacteurs ont fusionné, de sorte que si votre travail dépend du travail des autres utilisateurs (par exemple, les images partagées), il est disponible dans la nouvelle branche. Vous pouvez vérifier que votre travail précédent (et d’autres personnes) se trouve dans la succursale en extrayant la nouvelle branche...

    git checkout X2

...et vérifier le contenu. (La `checkout` commande met à jour les fichiers dans `C:\Users\<yourusername>\microsoft-graph-docs` l’état actuel de la branche x2.) Une fois que vous avez extrait la nouvelle branche, vous pouvez effectuer des mises à jour du contenu et les valider normalement. Toutefois, pour éviter de travailler par erreur dans la branche fusionnée (X), il est préférable de la supprimer (consultez la section suivante **Supprimer une branche**).

#### <a name="delete-a-branch"></a>Supprimer une branche

Une fois que vos modifications ont été fusionnées dans le référentiel principal, supprimez la branche utilisée car vous n’en avez plus besoin.  Tout travail supplémentaire doit être réalisé dans une nouvelle branche.  

#### <a name="to-delete-a-branch"></a>Pour supprimer une branche

1.  Dans l’invite de commandes git bash, tapez `git checkout master` . Ainsi, vous êtes sûr de ne pas vous trouver dans la branche à supprimer (ce qui est interdit).
2.  Ensuite, à l’invite de commandes, tapez `git branch -d <branch name>` . Cette opération supprime la branche sur votre ordinateur uniquement si elle a été correctement fusionnée avec le référentiel en amont. (Vous pouvez remplacer ce comportement par l’indicateur `–D`, mais vous devez d’abord être sûr de vouloir effectuer cette action.)
3.  Enfin, entrez `git push origin :<branch name>` à l’invite de commandes (un espace avant le signe deux-points et pas d’espace après celui-ci).  La branche est supprimée de votre bifurcation Github.  

Félicitations, vous avez contribué au projet.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Comment utiliser Markdown pour mettre en forme votre rubrique

### <a name="markdown"></a>Markdown

Tous les articles de ce référentiel utilisent Markdown. Vous trouverez une présentation complète (et la liste de toutes les syntaxes) à l’adresse [Daring Fireball-démarque].
 
## <a name="faq"></a>FAQ

### <a name="how-do-i-get-a-github-account"></a>Comment obtenir un compte GitHub ?

Remplissez le formulaire sur [Rejoindre GitHub](https://github.com/join) pour ouvrir un compte GitHub gratuit. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Où puis-je obtenir un Contrat de licence de contributeur ? 

Vous recevrez automatiquement une notification vous informant que vous devez signer le Contrat de licence de contributeur (CLA) si votre requête de tirage l’exige. 

En tant que membre de la communauté, **vous devez signer le Contrat de licence de contributeur (CLA) pour pouvoir contribuer largement à ce projet**. Vous ne devez remplir et soumettre la documentation qu’une seule fois. Lisez attentivement le document. Il faudra peut-être que votre employeur signe le document.

### <a name="what-happens-with-my-contributions"></a>Que se passe-t-il avec mes contributions ?

Lorsque vous envoyez vos modifications, via une demande de tirage, notre équipe sera informée et examinera votre requête de tirage. Vous recevrez des notifications sur votre demande de tirage à partir de GitHub ; vous pouvez également être notifié par une personne de notre équipe si vous avez besoin d’informations supplémentaires. Si votre demande de tirage est approuvée, nous mettrons à jour la documentation. Nous nous réservons le droit de modifier votre soumission pour des problèmes légaux, de style, de clarté ou autres.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Puis-je devenir un approbateur pour les demandes d’extraction GitHub de ce référentiel ?

Actuellement, nous n’autorisons pas les contributeurs externes à approuver les demandes d’extraction dans ce référentiel.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Comment puis-je obtenir une réponse à propos de ma demande de modification ?

Les requêtes de tirage sont généralement revues dans les 10 jours ouvrés.


## <a name="more-resources"></a>Autres ressources

* Pour en savoir plus sur la démarque, accédez au site du créateur de la démarque [Daring Fireball].
* Pour en savoir plus sur l’utilisation de git et GitHub, consultez [d']abord l’aide de github.

[GitHub Home]: http://github.com
[Aide de GitHub]: http://help.github.com/
[Configurer Git]: https://help.github.com/articles/set-up-git/
[Daring Fireball-démarque]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
