> [!TIP]
> Si vous testez votre complément dans plusieurs environnements (par exemple, en cours de développement, de mise en œuvre, de démonstration, etc.), nous vous recommandons de conserver un fichier manifeste XML différent pour chaque environnement. Dans chaque fichier manifeste, vous pouvez :
> - Spécifiez les URL qui correspondent à l’environnement.
> - Personnalisez les valeurs `DisplayName` de métadonnées `Resources` telles que et les étiquettes au sein pour indiquer l’environnement, afin que les utilisateurs finals puissent identifier l’environnement correspondant d’un complément versions test chargées. 
> - Personnalisez les fonctions `namespace` personnalisées pour indiquer l’environnement, si votre complément définit des fonctions personnalisées.
> 
> En suivant ce guide, vous allez rationaliser le processus de test et éviter les problèmes qui seraient normalement survenus lorsqu’un complément est simultanément versions test chargées pour plusieurs environnements.