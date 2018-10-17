# <a name="contribute-to-this-documentation"></a>Contribuer à cette documentation

Nous vous remercions de l’intérêt que vous portez à notre documentation.

* [Méthodes de contribution](#ways-to-contribute)
* [Contribution à l’aide de GitHub](#contribute-using-github)
* [Contribution à l’aide de Git](#contribute-using-git)
* [Utilisation de Markdown pour mettre en forme votre rubrique](#how-to-use-markdown-to-format-your-topic)
* [FAQ](#faq)
* [Autres ressources](#more-resources)

## <a name="ways-to-contribute"></a>Méthodes de contribution

Voici comment vous pouvez contribuer à la documentation :

* Pour apporter des modifications mineures à un article, reportez-vous à la section [Contribution à l’aide de GitHub](#contribute-using-github).
* Pour apporter des modifications importantes ou relatives au code, reportez-vous à la section [Contribution à l’aide de Git](#contribute-using-git).
* Signalez les bogues de documentation via Problèmes GitHub
* Demandez la nouvelle documentation sur le site [Plateforme UserVoice pour les développeurs Office](http://officespdev.uservoice.com).

## <a name="contribute-using-github"></a>Contribution à l’aide de GitHub

Utilisez GitHub pour contribuer à cette documentation sans avoir à cloner le référentiel sur votre bureau. Il s’agit de la manière la plus simple de créer une demande de tirage (pull request) dans ce référentiel. Utilisez cette méthode pour apporter des changements mineurs qui n’impliquent pas de modifier le code. 

**Remarque** : cette méthode vous permet de contribuer à un article à la fois.

### <a name="to-contribute-using-github"></a>Contribution à l’aide de GitHub

1. Recherchez l’article auquel vous souhaitez contribuer sur GitHub.
2. Une fois que vous avez accédé à l’article dans GitHub, connectez-vous à GitHub (obtenir un compte gratuit [Rejoindre GitHub](https://github.com/join)).
3. Choisissez l’**icône en forme de crayon** (modifier le fichier dans la bifurcation de ce projet) et apportez vos modifications dans la fenêtre **<>Modifier le fichier**. 
4. Faites défiler vers le bas, puis entrez une description.
5. Choisissez **Proposer une modification de fichier**>**Créer une demande de tirage (pull request)**.

Vous avez désormais envoyé une demande de tirage. Les demandes de tirage sont généralement examinées dans les 10 jours ouvrés. 


## <a name="contribute-using-git"></a>Contribution à l’aide de Git

Utilisez Git pour apporter des modifications substantielles, telles que :

* Contribution au code.
* Contribution aux modifications qui ont une incidence sur la signification.
* Contribution à des changements importants apportés au texte.
* Ajout de nouvelles rubriques.

### <a name="to-contribute-using-git"></a>Contribution à l’aide de Git

1. Si vous n’avez pas de compte GitHub, configurez-en un sur [GitHub](https://github.com/join). 
2. Une fois que vous disposez d’un compte, installez Git sur votre ordinateur. Suivez les étapes indiquées dans le [didacticiel sur la configuration de Git].
3. Pour envoyer une demande de tirage à l’aide de Git, suivez les étapes indiquées dans la section [Utilisation de GitHub, de Git et du référentiel](#use-github-git-and-this-repository).
4. Vous serez invité à signer le Contrat de licence de contributeur si vous êtes :

    * membre du groupe Technologies ouvertes Microsoft ;
    * un contributeur qui ne travaille pas pour Microsoft.

En tant que membre de la communauté, vous devez signer le Contrat de licence de contributeur (CLA) pour pouvoir apporter des contributions conséquentes à un projet. Vous ne devez remplir et soumettre la documentation qu’une seule fois. Lisez attentivement le document. Il faudra peut-être que votre employeur signe le document.

En signant ce contrat, vous n’obtenez pas le droit de valider le référentiel principal. Cela signifie seulement que les équipes Publication du contenu de développeur Office et Développeurs Office pourront examiner et approuver vos contributions. Votre nom apparaîtra comme auteur de vos contributions.

Les demandes de tirage sont généralement examinées dans les 10 jours ouvrés.

## <a name="use-github-git-and-this-repository"></a>Utilisation de GitHub, de Git et du référentiel

**Remarque :** La plupart des informations contenues dans cette section figurent dans les articles de l’[aide GitHub].  Si vous connaissez Git et GitHub, passez à la section **Contribution et modification du contenu** pour découvrir les éléments du flux de code/contenu de ce référentiel.

### <a name="to-set-up-your-fork-of-the-repository"></a>Configuration de votre bifurcation de référentiel

1.  Vous devez avoir configuré un compte GitHub pour pouvoir contribuer à ce projet. Si ce n'est pas déjà fait, accédez à [GitHub](https://github.com/join) et configurez le compte maintenant.
2.  Installez Git sur votre ordinateur. Suivez les étapes indiquées dans le [didacticiel sur la configuration de Git].
3.  Créez votre propre bifurcation du référentiel. Pour cela, cliquez sur le bouton **Dupliquer (Fork)** en haut de la page.
4.  Copiez la duplication sur votre ordinateur. Pour ce faire, ouvrez Git Bash. À l’invite de commandes, entrez les informations suivantes :

        git clone https://github.com/<your user name>/<repo name>.git

    Créez ensuite une référence vers le référentiel racine en entrant les commandes suivantes  :

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Félicitations ! Votre référentiel est maintenant configuré. Vous n’aurez plus besoin de répéter ces étapes.

### <a name="contribute-and-edit-content"></a>Contribution et modification de contenu

Pour contribuer de la manière la plus simple possible, suivez les étapes suivantes.

#### <a name="to-contribute-and-edit-content"></a>Contribution et modification de contenu

1. Créez une branche.
2. Ajoutez du nouveau contenu ou modifiez le contenu existant.
3. Envoyez une demande de tirage au référentiel principal.
4. Supprimez la branche.

**Important** Attribuez un concept/article par branche pour simplifier le flux de travail et réduire les risques de conflits de fusion. Le contenu approprié à une nouvelle branche inclut les éléments suivants :

* Nouvel article.
* Modifications liées à l’orthographe et à la grammaire.
* Changement de mise en forme unique dans plusieurs articles (par exemple, insérer un copyright dans le pied de page).

#### <a name="to-create-a-new-branch"></a>Création d’une branche

1.  Ouvrez Git Bash.
2.  À l’invite de commande Git Bash, saisissez `git pull upstream master:<new branch name>`. Une branche est créée localement et copiée à partir de la dernière branche maître OfficeDev.
3.  À l’invite de commande Git Bash, saisissez `git push origin <new branch name>`. Ceci signale la nouvelle branche à GitHub. La nouvelle branche doit maintenant apparaître dans votre bifurcation du référentiel sur GitHub.
4.  À l’invite de commande Git Bash, saisissez `git checkout <new branch name>` pour basculer sur la nouvelle branche.

#### <a name="add-new-content-or-edit-existing-content"></a>Ajout de nouveau contenu ou modification de contenu existant

Utilisez l’Explorateur de fichiers pour accéder au référentiel sur votre ordinateur. Les fichiers du référentiel se trouvent dans `C:\Users\<yourusername>\<repo name>`.

Pour modifier des fichiers, ouvrez-les dans l’éditeur de votre choix et modifiez-les. Pour créer un fichier, utilisez l’éditeur de votre choix et enregistrez le nouveau fichier à l’emplacement approprié dans votre copie locale du référentiel. Lorsque vous travaillez, enregistrez votre travail régulièrement.

Les fichiers présents dans `C:\Users\<yourusername>\<repo name>` constituent une copie de travail de la nouvelle branche créée dans le référentiel local. Si vous effectuez des modifications dans ce dossier, le référentiel local ne sera pas altéré tant que vous ne validez pas les modifications. Pour valider une modification dans le référentiel local, tapez les commandes suivantes dans GitBash :

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

La commande `add` ajoute vos modifications dans une zone de transit pour qu’elles soient validées dans le référentiel. Le point situé après la commande `add` indique que tous les fichiers ajoutés ou modifiés doivent être placés dans la zone de transit, en vérifiant les sous-dossiers de manière récursive. (Si vous ne souhaitez pas valider toutes les modifications, vous pouvez ajouter des fichiers spécifiques. Vous pouvez également annuler une validation. Pour plus d’informations, entrez `git add -help` ou `git status`.)

La commande `commit` applique les modifications en attente dans le référentiel. Le commutateur « `-m` » signifie que vous fournissez le commentaire de validation dans la ligne de commande. Les commutateurs « -v » et « -a » peuvent être omis. Le commutateur « -v » est détaillé à partir de la commande, tandis que le commutateur « -a » répète l’action effectuée avec la commande « add ».

Vous pouvez valider des changements plusieurs fois lorsque vous travaillez, ou vous pouvez tout valider une fois que vous avez terminé.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Envoi d’une demande de tirage au référentiel principal

Quand vous avez terminé votre travail et que vous êtes prêt à le fusionner dans le référentiel principal, suivez les étapes suivantes.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Envoi d’une requête de tirage au référentiel principal.

1.  Dans l’invite de commande Git Bash, saisissez `git push origin <new branch name>`. Dans le référentiel local, `origin` fait référence au référentiel GitHub à partir duquel vous avez cloné le référentiel local. Cette commande affiche l’état actuel de votre nouvelle branche, y compris toutes les validations effectuées au cours des étapes précédentes dans votre bifurcation GitHub.
2.  Sur le site GitHub, accédez à la nouvelle branche dans votre bifurcation.
3.  Sélectionnez le bouton **Demande de tirage** situé en haut de la page.
4.  Vérifiez que la branche de base est `OfficeDev/<repo name>@master` et que la branche de tête est `<your username>/<repo name>@<branch name>`.
5.  Sélectionnez le bouton **Mettre à jour la plage de validation**.
6.  Ajoutez un titre à votre demande de tirage et décrivez toutes les modifications que vous apportez.
7.  Envoyez la requête de tirage.

Un des administrateurs du site traitera votre demande de tirage. Votre demande de tirage apparaîtra sur le site OfficeDev/<repo name> sous Problèmes. Lorsque votre la demande de tirage est acceptée, le problème est résolu.

#### <a name="create-a-new-branch-after-merge"></a>Créer une branche après fusion

Après la fusion d’une branche (par exemple, quand votre demande de tirage est acceptée), cessez de travailler dans la branche locale. Ceci peut entraîner la fusion des conflits si vous envoyez une autre demande de tirage. Pour effectuer une autre mise à jour, créez une branche locale à partir de la branche fusionnée en amont, puis supprimez la branche locale d’origine.

Par exemple, si votre branche locale X a été fusionnée dans la branche maître OfficeDev/microsoft-graph-docs et que vous voulez effectuer d’autres mises à jour dans le contenu fusionné. Créez une branche locale, X2, à partir de la branche maître OfficeDev/microsoft-graph-docs. Pour cela, ouvrez GitBash et exécutez les commandes suivantes :

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Vous possédez désormais des copies locales (dans une nouvelle branche locale) du travail que vous avez effectué dans la branche X. La branche X2 contient également tout le travail que les autres auteurs ont fusionné. Ainsi, si votre travail dépend du travail d’autres personnes (par exemple, des images partagées), celui-ci est disponible dans la nouvelle branche. Pour vérifier que le travail que vous avez effectué (ainsi que celui d’autres personnes) se trouve dans la branche, vous pouvez extraire la nouvelle branche...

    git checkout X2

...et consulter le contenu. (La commande `checkout` applique les mises à jour des fichiers dans `C:\Users\<yourusername>\microsoft-graph-docs` à l’état actuel de la branche X2.) Quand vous avez extrait la nouvelle branche, vous pouvez mettre à jour le contenu et valider les mises à jour. Toutefois, pour éviter de travailler par erreur dans la branche fusionnée (X), il est préférable de la supprimer (consultez la section suivante **Supprimer une branche**).

#### <a name="delete-a-branch"></a>Supprimer une branche

Lorsque vos modifications sont fusionnées dans le référentiel principal, supprimez la branche utilisée dont vous n’avez plus besoin.  Aucun travail supplémentaire ne doit être effectué dans une nouvelle branche.  

#### <a name="to-delete-a-branch"></a>Suppression d’une branche

1.  Dans l’invite de commande Git Bash, saisissez `git checkout master`. Ainsi, vous êtes sûr de ne pas vous trouver dans la branche à supprimer (ce qui est interdit).
2.  À l’invite de commande, saisissez `git branch -d <branch name>`. La branche est supprimée de votre ordinateur uniquement si elle a été fusionnée en amont dans le référentiel. (Vous pouvez remplacer ce comportement par l’indicateur `–D`, mais vous devez d’abord être sûr de vouloir effectuer cette action.)
3.  Enfin, entrez `git push origin :<branch name>` à l’invite de commandes (un espace avant le signe deux-points et pas d’espace après celui-ci).  La branche est supprimée de votre bifurcation Github.  

Félicitations, vous avez apporté une contribution au projet !

## <a name="how-to-use-markdown-to-format-your-topic"></a>Utilisation de Markdown pour mettre en forme votre rubrique

### <a name="markdown"></a>Markdown

Tous les articles de ce référentiel utilisent Markdown. Une présentation complète (et la description de toutes les syntaxes) est disponible sur la page [Daring Fireball - Markdown].
 
## <a name="faq"></a>FAQ

### <a name="how-do-i-get-a-github-account"></a>Comment obtenir un compte GitHub ?

Remplissez le formulaire sur [Rejoindre GitHub](https://github.com/join) pour ouvrir un compte GitHub gratuit. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Où puis-je obtenir un Contrat de licence de contributeur ? 

Vous recevrez automatiquement une notification vous informant que vous devez signer le Contrat de licence de contributeur (CLA) si votre requête de tirage l’exige. 

En tant que membre de la communauté, **vous devez signer le Contrat de licence de contributeur (CLA) pour pouvoir apporter des contributions conséquentes à un projet**. Vous ne devez remplir et soumettre la documentation qu’une seule fois. Lisez attentivement le document. Il faudra peut-être que votre employeur signe le document.

### <a name="what-happens-with-my-contributions"></a>Que se produit-t-il avec mes contributions ?

Lorsque vous envoyez vos modifications, via une demande de tirage, notre équipe est informée et examine votre demande. Vous recevrez des notifications relatives à votre demande de tirage de la part de GitHub. Vous pouvez également être contacté par une personne de notre équipe si nous avons besoin de plus d’informations. Si votre demande est approuvée, nous mettrons à jour la documentation. Nous nous réservons le droit de modifier votre envoi en cas de problèmes juridiques, de style, de clarté ou autres.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Puis-je devenir un approbateur pour les demandes de tirage GitHub de ce référentiel ?

Actuellement, nous n’autorisons pas les contributeurs externes à approuver les demandes de tirage dans ce référentiel.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Quel est le délai d’attente avant d’obtenir une réponse à propos de ma demande de modification ?

Les demandes de tirage sont généralement examinées dans les 10 jours ouvrés.


## <a name="more-resources"></a>Autres ressources

* Pour en savoir plus sur Markdown, accédez au site du créateur de Markdown, [Daring Fireball].
* Pour en savoir plus sur l’utilisation de Git et GitHub, commencez par consulter la [l’aide GitHub].

[GitHub Home]: http://github.com
[Aide de GitHub]: http://help.github.com/
[Configuration de Git]: https://help.github.com/articles/set-up-git/
[Daring Fireball - Markdown]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
