# <a name="contribute-to-this-documentation"></a>Contribuer à cette documentation

Nous vous remercions de l’intérêt que vous portez à notre documentation.

* [Méthodes de collaboration](#ways-to-contribute)
* [Contribuer à l’aide de GitHub](#contribute-using-github)
* [Contribuer à l’aide de Git](#contribute-using-git)
* [Utilisation de Markdown pour mettre en forme votre rubrique](#how-to-use-markdown-to-format-your-topic)
* [FAQ](#faq)
* [Autres ressources](#more-resources)

## <a name="ways-to-contribute"></a>Méthodes de collaboration

Voici quelques façons de contribuer à cette documentation :

* Pour apporter de petites modifications à un article, [contribuez à l’aide GitHub](#contribute-using-github).
* Pour apporter des modifications importantes ou des modifications qui impliquent du code, [contribuez à l’aide de Git](#contribute-using-git).
* Signalez les bogues de **documentation** en allant à la section Commentaires au bas de l’article concerné, puis en sélectionnant Cette **page** pour créer un GitHub problème. Si ce n’est pas disponible, créez un problème directement sur [GitHub](https://github.com/OfficeDev/office-js-docs-pr/issues).
* Demandez une nouvelle documentation avec [GitHub problèmes.](https://github.com/OfficeDev/office-js-docs-pr/issues)

## <a name="contribute-using-github"></a>Contribuer à l’aide de GitHub

Utilisez GitHub pour contribuer à cette documentation sans avoir à cloner le repo sur votre bureau. Il s’agit du moyen le plus simple de créer une demande de tirage dans ce référentiel. Utilisez cette méthode pour apporter une modification mineure qui n’implique pas de modifications de code.

**Remarque**: l’utilisation de cette méthode vous permet de contribuer à un article à la fois.

### <a name="to-contribute-using-github"></a>Pour contribuer à l’utilisation GitHub

1. Recherchez l’article sur GitHub.
2. Une fois que vous êtes sur l’article GitHub, connectez-vous à GitHub (obtenez un compte gratuit [rejoindre GitHub](https://github.com/join)).
3. Choisissez **l’icône de crayon** (modifiez le fichier dans votre bifurcation de ce projet) et a apporté vos modifications dans la **fenêtre<>modifier le** fichier.
4. Faites défiler vers le bas et entrez une description.
5. Choose **Propose file change** Create pull > **request**.

Vous avez maintenant envoyé une demande de pull. Les demandes de pull sont généralement examinées dans les 10 jours ou moins.


## <a name="contribute-using-git"></a>Contribuer à l’aide de Git

Utilisez Git pour apporter des modifications importantes, telles que :

* Code de contribution.
* Modifications qui ont une incidence sur la signification.
* Contribution de modifications importantes au texte.
* Ajout de nouvelles rubriques.

### <a name="to-contribute-using-git"></a>Pour contribuer à l’aide de Git

1. Si vous n’avez pas de compte GitHub, définissez-en un sur [GitHub](https://github.com/join).
2. Une fois que vous avez un compte, installez Git sur votre ordinateur. Suivez les étapes du [didacticiel Configurer Git.]
3. Pour envoyer une requête de pull à l’aide de Git, suivez les étapes dans [Utiliser GitHub, Git et ce référentiel.](#use-github-git-and-this-repository)
4. Vous serez invité à signer le contrat de licence du collaborateur si vous êtes :

    * Membre du groupe Microsoft Open Technologies.
    * Collaborateur qui ne travaille pas pour Microsoft.

En tant que membre de la communauté, vous devez signer le contrat de licence de contribution (CLA) avant de pouvoir contribuer à des soumissions importantes à un projet. Vous ne devez remplir et envoyer la documentation qu’une seule fois. Lisez attentivement le document. Il faudra peut-être que votre employeur signe le document.

La signature du CLA ne vous accorde pas le droit de valider dans le référentiel principal, mais cela signifie que les équipes de publication de contenu du développeur Office et du développeur Office pourront examiner et approuver vos contributions. Vous êtes crédité pour vos soumissions.

Les demandes de pull sont généralement examinées dans les 10 jours ou moins.

## <a name="use-github-git-and-this-repository"></a>Utilisation de GitHub, de Git et du référentiel

**Remarque**: la plupart des informations de cette section sont disponibles dans [GitHub’aide.]  Si vous connaissez Git et GitHub, passez à la **section** Contribuer et modifiez le contenu pour connaître les spécificités du flux de code/contenu de ce référentiel.

### <a name="to-set-up-your-fork-of-the-repository"></a>Pour configurer votre bifurcation du référentiel

1. Configurez un compte GitHub pour pouvoir contribuer à ce projet. Si vous ne l’avez pas fait, [GitHub](https://github.com/join) et faites-le maintenant.
2. Installez Git sur votre ordinateur. Suivez les étapes du [didacticiel Configurer Git.]
3. Créez votre propre bifurcation du référentiel. Pour ce faire, en haut de la page, choisissez le **bouton Bifurcation.**
4. Copiez votre bifurcation sur votre ordinateur. Pour ce faire, ouvrez Git Bash. À l’invite de commandes, entrez les informations suivantes :

        git clone https://github.com/<your user name>/<repo name>.git

    Ensuite, créez une référence au référentiel racine en entrant ces commandes.

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Félicitations ! Votre référentiel est maintenant configuré. Vous n’aurez plus besoin d’effectuer ces étapes.

### <a name="contribute-and-edit-content"></a>Contribution et modification de contenu

Pour rendre le processus de contribution aussi transparent que possible, suivez ces étapes.

#### <a name="to-contribute-and-edit-content"></a>Pour contribuer et modifier du contenu

1. Créez une branche.
2. Ajoutez du nouveau contenu ou modifiez le contenu existant.
3. Envoyez une requête de tirage au référentiel principal.
4. Supprimez la branche.

**Important** Limitez chaque branche à un seul concept/article pour simplifier le flux de travail et réduire le risque de conflits de fusion. Le contenu approprié pour une nouvelle branche inclut :

* Un nouvel article.
* Modifications orthographiques et grammaticales.
* Application d’une seule modification de mise en forme sur un grand ensemble d’articles (par exemple, application d’un nouveau pied de copyright).

#### <a name="to-create-a-new-branch"></a>Pour créer une branche

1. Ouvrez Git Bash.
2. À l’invite de commandes Git Bash, tapez `git pull upstream master:<new branch name>` . Cela crée une nouvelle branche localement qui est copiée à partir de la dernière branche maître OfficeDev.
3. À l’invite de commandes Git Bash, tapez `git push origin <new branch name>` . Cette alerte GitHub la nouvelle branche. La nouvelle branche doit maintenant apparaître dans votre bifurcation du référentiel de GitHub.
4. À l’invite de commandes Git Bash, `git checkout <new branch name>` tapez pour basculer vers votre nouvelle branche.

#### <a name="add-new-content-or-edit-existing-content"></a>Ajout de nouveau contenu ou modification de contenu existant

Vous accédez au référentiel sur votre ordinateur à l’aide de l’Explorateur de fichiers. Les fichiers de référentiel sont dans `C:\Users\<yourusername>\<repo name>` .

Pour modifier des fichiers, ouvrez-les dans un éditeur de votre choix et modifiez-les. Pour créer un fichier, utilisez l’éditeur de votre choix et enregistrez le nouveau fichier à l’emplacement approprié dans votre copie locale du référentiel. Tout en travaillant, enregistrez fréquemment votre travail.

Les fichiers qu’ils contiennent sont une copie de travail de la nouvelle branche que `C:\Users\<yourusername>\<repo name>` vous avez créée dans votre référentiel local. Si vous effectuez des modifications dans ce dossier, le référentiel local ne sera pas altéré tant que vous ne validez pas les modifications. Pour valider une modification dans le référentiel local, tapez les commandes suivantes dans GitBash.

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

La commande `add` ajoute vos modifications dans une zone de transit pour qu’elles soient validées dans le référentiel. Le point après la commande spécifie que vous souhaitez mettre en avant tous les fichiers que vous avez ajoutés ou modifiés, en vérifiant les sous-dossiers `add` de manière récursive. (Si vous ne souhaitez pas valider toutes les modifications, vous pouvez ajouter des fichiers spécifiques. Vous pouvez également annuler une validation. Pour obtenir de l’aide, entrez `git add -help` ou `git status`.)

La commande `commit` applique les modifications en attente dans le référentiel. Le commutateur `-m` signifie que vous fournissez le commentaire de validation dans la ligne de commande. Les commutateurs -v et -a peuvent être omis. Le commutateur -v est pour la sortie détaillée de la commande, et -a fait ce que vous avez déjà fait avec la commande Ajouter.

Vous pouvez valider plusieurs fois pendant votre travail ou une seule fois lorsque vous avez terminé.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Envoyer une requête de tirage au référentiel principal

Lorsque vous avez terminé votre travail et que vous êtes prêt à le fusionner dans le référentiel principal, suivez ces étapes.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Pour envoyer une demande de pull au référentiel principal

1. Dans l’invite de commandes Git Bash, tapez `git push origin <new branch name>` . Dans le référentiel local, `origin` fait référence au référentiel GitHub à partir duquel vous avez cloné le référentiel local. Cette commande affiche l’état actuel de votre nouvelle branche, y compris toutes les validations effectuées au cours des étapes précédentes dans votre bifurcation GitHub.
2. Sur le site GitHub, accédez à la nouvelle branche dans votre bifurcation.
3. Sélectionnez **le bouton Tirer la** demande en haut de la page.
4. Vérifiez que la branche base `OfficeDev/<repo name>@master` est et que la branche Head est `<your username>/<repo name>@<branch name>` .
5. Sélectionnez le **bouton Mettre à jour la plage de validation.**
6. Ajoutez un titre à votre requête de pull et décrivez toutes les modifications que vous a faites.
7. Envoyez la requête de tirage.

L’un des administrateurs de site traitera votre demande de pull. Votre requête de pull s’surface sur le site OfficeDev/sous <repo name> Problèmes. Lorsque la demande de pull est acceptée, le problème est résolu.

#### <a name="create-a-new-branch-after-merge"></a>Créer une branche après fusion

Une fois qu’une branche a été fusionnée (c’est-à-dire que votre demande de pull est acceptée), ne continuez pas à travailler dans cette branche locale. Cela peut entraîner des conflits de fusion si vous envoyez une autre demande de pull. Pour faire une autre mise à jour, créez une nouvelle branche locale à partir de la branche fusionnée en amont correctement, puis supprimez votre branche locale initiale.

Par exemple, si votre branche X locale a été fusionnée avec succès dans la branche maître OfficeDev/microsoft-graph-docs et que vous souhaitez apporter des mises à jour supplémentaires au contenu qui a été fusionné. Créez une branche locale, X2, à partir de la branche maître OfficeDev/microsoft-graph-docs. Pour ce faire, ouvrez GitBash et exécutez les commandes suivantes.

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Vous avez maintenant des copies locales (dans une nouvelle branche locale) du travail que vous avez soumis dans la branche X. La branche X2 contient également tout le travail que d’autres rédacteurs ont fusionné. Ainsi, si votre travail dépend du travail d’autres personnes (par exemple, des images partagées), il est disponible dans la nouvelle branche. Vous pouvez vérifier que votre travail précédent (et celui d’autres personnes) se trouve dans la branche en vérifiant la nouvelle succursale...

    git checkout X2

...et vérifier le contenu. (La `checkout` commande met à jour les fichiers à `C:\Users\<yourusername>\microsoft-graph-docs` l’état actuel de la branche X2.) Une fois que vous avez vérifié la nouvelle branche, vous pouvez mettre à jour le contenu et les valider comme d’habitude. Toutefois, pour éviter de travailler par erreur dans la branche fusionnée (X), il est préférable de la supprimer (consultez la section suivante **Supprimer une branche**).

#### <a name="delete-a-branch"></a>Supprimer une branche

Une fois vos modifications correctement fusionnées dans le référentiel principal, supprimez la branche que vous avez utilisée, car vous n’en avez plus besoin.  Tout travail supplémentaire doit être effectué dans une nouvelle branche.  

#### <a name="to-delete-a-branch"></a>Pour supprimer une branche

1. Dans l’invite de commandes Git Bash, tapez `git checkout master` . Ainsi, vous êtes sûr de ne pas vous trouver dans la branche à supprimer (ce qui est interdit).
2. Ensuite, à l’invite de commandes, tapez `git branch -d <branch name>` . Cela supprime la branche sur votre ordinateur uniquement si elle a été fusionnée avec succès dans le référentiel en amont. (Vous pouvez remplacer ce comportement par l’indicateur `–D`, mais vous devez d’abord être sûr de vouloir effectuer cette action.)
3. Enfin, entrez `git push origin :<branch name>` à l’invite de commandes (un espace avant le signe deux-points et pas d’espace après celui-ci).  La branche est supprimée de votre bifurcation Github.  

Félicitations, vous avez contribué au projet !

## <a name="how-to-use-markdown-to-format-your-topic"></a>Comment utiliser Markdown pour mettre en forme votre rubrique

### <a name="markdown"></a>Markdown

Tous les articles de ce référentiel utilisent Markdown. Vous pouvez trouver une introduction complète (et la liste de toutes les syntaxes) sur le site [Fireball de l’insérez - Markdown].

## <a name="faq"></a>FAQ

### <a name="how-do-i-get-a-github-account"></a>Comment obtenir un compte GitHub ?

Remplissez le formulaire sur [Rejoindre GitHub](https://github.com/join) pour ouvrir un compte GitHub gratuit.

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Où puis-je obtenir un Contrat de licence de contributeur ?

Vous recevrez automatiquement une notification vous informant que vous devez signer le Contrat de licence de contributeur (CLA) si votre requête de tirage l’exige.

En tant que membre de la communauté, **vous devez signer le Contrat de licence de contributeur (CLA) pour pouvoir contribuer largement à ce projet**. Vous ne devez remplir et soumettre la documentation qu’une seule fois. Lisez attentivement le document. Il faudra peut-être que votre employeur signe le document.

### <a name="what-happens-with-my-contributions"></a>Que se passe-t-il avec mes contributions ?

Lorsque vous soumettez vos modifications, via une demande de pull, notre équipe est avertie et examine votre demande de pull. Vous recevrez des notifications concernant votre demande de pull de GitHub ; Vous pouvez également être averti par une personne de notre équipe si nous avons besoin d’informations supplémentaires. Si votre requête de pull est approuvée, nous allons mettre à jour la documentation. Nous nous réservons le droit de modifier votre soumission pour des raisons juridiques, de style, de clarté ou d’autres problèmes.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Puis-je devenir un approuveur pour les demandes de GitHub de ce référentiel ?

Actuellement, nous ne permettons pas aux collaborateurs externes d’approuver les demandes de pull dans ce référentiel.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>À combien de temps vais-je obtenir une réponse concernant ma demande de modification ?

Les demandes de pull sont généralement examinées dans les 10 jours ou moins.


## <a name="more-resources"></a>Plus de ressources

* Pour en savoir plus sur Markdown, allez sur le site du créateur Markdown [« Fireball ».]
* Pour en savoir plus sur l’utilisation de Git et GitHub, consultez d’abord [l’aide GitHub.]

[GitHub Home]: http://github.com
[Aide de GitHub]: http://help.github.com/
[Configurer Git]: https://help.github.com/articles/set-up-git/
[Fireball de puissance - Markdown]: http://daringfireball.net/projects/markdown/
[Fireball de puissance]: http://daringfireball.net/
