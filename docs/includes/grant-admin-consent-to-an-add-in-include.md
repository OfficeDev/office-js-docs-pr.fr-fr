
> [!NOTE]
> Cette procédure est uniquement nécessaire quand vous développez le complément. Lorsque votre application de production est déployée sur AppSource ou un catalogue d’applications, les utilisateurs l’utilisent individuellement ou un administrateur consent à l’organisation lors de l’installation.

Effectuez cette procédure *une fois* que vous avez inscrit [le add-in.](../develop/register-sso-add-in-aad-v2.md) (Si vous avez terminé cette procédure et que l’onglet **Autorisations d’API** de la page **$ADD-IN-NAME$** est ouvert dans votre  navigateur, vous pouvez choisir le bouton Accorder le consentement administrateur pour [nom du **client],** puis sélectionner Oui pour la confirmation qui s’affiche. Ignorez le reste de cette procédure.)

1. Accédez à la page [Portail Azure - Inscriptions d’applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour afficher l’inscription de votre application.

1. Connectez-vous avec ***les informations d’identification*** d’administrateur Microsoft 365 location. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez l’application avec le **nom complet $ADD-IN-NAME$**.

1. Sur la page **$ADD-IN-NAME$,** sélectionnez les **autorisations d’API,** puis, sous la **section** Accorder le consentement, sélectionnez le bouton Accorder le consentement administrateur pour **[nom** du client]. Sélectionnez **Oui** pour la confirmation qui s’affiche.

> [!NOTE]
> Nous vous recommandons d’utiliser cette procédure comme meilleure pratique si vous utilisez un client O365 développeur. Toutefois, si vous préférez, il est possible de recharger une version de chargement d’une version de l' utilisateur en cours de développement et d’inviter l’utilisateur avec un formulaire de consentement. Pour plus d’informations, voir [Sideload on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) and [Sideload on Office sur le Web](../testing/sideload-office-add-ins-for-testing.md).
