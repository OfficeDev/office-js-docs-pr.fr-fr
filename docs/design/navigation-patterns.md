# <a name="navigation-patterns"></a>Modèles de navigation

Les principales fonctionnalités d’un complément sont accessibles via des types de commandes spécifiques et une zone d’écran limitée. Il est important que la navigation soit intuitive, qu'elle fournisse un contexte et permette à l‘utilisateur de se déplacer facilement dans toute l'étendue du complément.

## <a name="best-practices"></a>Meilleures pratiques

| À faire    | À ne pas faire |
| :---- | :---- |
| Assurez-vous que l’utilisateur dispose d’une option de navigation clairement visible. | Ne compliquez pas le processus de navigation en utilisant une interface utilisateur non standard.
| Utilisez les composants suivants, le cas échéant, pour permettre aux utilisateurs de naviguer sur l'étendue de votre complément. | L’utilisateur doit comprendre sa place ou son contexte actuel dans le complément, c’est pourquoi vous ne devez pas lui compliquer la tâche



## <a name="command-bar"></a>Barre de commandes

La barre de commandes est une surface qui héberge des commandes qui agissent sur le contenu de la fenêtre, du panneau ou de la région parent au-dessus de laquelle elle réside. Les caractéristiques facultatives comprennent le point d’accès au menu hamburger, la recherche et des commandes latérales.

![Commandes — Spécifications pour le volet des tâches de bureau](../images/add-in-command-bar.png)



## <a name="tab-bar"></a>Barre d’onglets

Barre d’onglets — Affiche la navigation à l’aide de boutons avec du texte et des icônes verticalement empilés. Utiliser la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.

![Barre d’onglets — Spécifications du volet des tâches de bureau](../images/add-in-tab-bar.png)


## <a name="back-button"></a>Retour

Le bouton Précédent permet aux utilisateurs de revenir au stade intial après avoir exécuté une action de navigation détaillée. Utilisez ce modèle pour vous assurer que les utilisateurs suivent une série d’étapes ordonnées.  

![Bouton Précédent — Spécifications du volet des tâches de bureau](../images/add-in-back-button.png)
