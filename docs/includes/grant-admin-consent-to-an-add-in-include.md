
> [!NOTE]
> Cette procédure est uniquement nécessaire quand vous développez le complément. Lorsque votre complément de production est déployé dans AppSource ou dans un catalogue d’applications, les utilisateurs l’approuvent individuellement ou un administrateur consentira à l’Organisation lors de l’installation.

Exécutez cette procédure *après* avoir [enregistré le complément](../develop/register-sso-add-in-aad-v2.md). (Si vous venez d’effectuer cette procédure et que l’onglet Autorisations de l' **API** de la page **$Add-in-Name $** est ouvert dans votre navigateur, vous pouvez choisir le bouton **accorder le consentement de l’administrateur pour [nom du client]** , puis sélectionner **Oui** pour la confirmation qui s’affiche. Ignorez le reste de cette procédure.)

1. Accédez à la page [Azure portal-inscriptions aux applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour afficher l’inscription de votre application.

1. Connectez-vous avec les informations d’identification d' ***administrateur*** à votre location Microsoft 365. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez l’application dont le nom d’affichage est **$Add-in-Name $**.

1. Sur la page **$Add-in-Name $** , sélectionnez **autorisations d’API** puis, sous la section **consentement de subvention** , sélectionnez le bouton **accorder le consentement de l’administrateur pour le bouton [nom du client]** . Sélectionnez **Oui** pour confirmer l’affichage.

> [!NOTE]
> Nous vous recommandons d’utiliser cette procédure en tant que meilleure pratique si vous utilisez un client O365 de développeur. Toutefois, si vous préférez, vous pouvez chargement un complément d’authentification unique en cours de développement et inviter l’utilisateur à fournir un formulaire de consentement. Pour plus d’informations, reportez-vous à [chargement sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) et [chargement sur le Web](../testing/sideload-office-add-ins-for-testing.md).
