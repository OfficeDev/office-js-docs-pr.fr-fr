> [!TIP]
> Si vous allez tester votre complément dans plusieurs environnements (par exemple, dans le développement, la gestion intermédiaire, la version de démonstration, etc.), nous vous recommandons de maintenir un autre fichier manifeste XML pour chaque environnement. Dans chaque fichier manifeste, vous pouvez :
> - Spécifier les URL qui correspondent à l’environnement.
> - Personnaliser des valeurs de métadonnée telles que `DisplayName` et les étiquettes dans `Resources` pour indiquer l’environnement pour que les utilisateurs finaux puissent identifier un environnement correspondant du complément chargé indépendamment. 
> - Personnaliser les fonctions `namespace` personnalisées pour indiquer l’environnement si votre complément définit des fonctions personnalisées.
> 
> En suivant ces conseils, vous simplifiez le processus de test et éviter des problèmes qui se produisent lorsqu’un complément est chargé indépendamment en même temps dans de nombreux environnements.