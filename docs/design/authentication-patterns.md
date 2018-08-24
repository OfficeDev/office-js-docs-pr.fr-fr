# <a name="authentication-patterns"></a>Modèles d'authentification

Les compléments peuvent exiger que les utilisateurs se connectent ou s'inscrivent afin d'accéder aux fonctions et fonctionnalités. Les zones de saisie du nom d'utilisateur et du mot de passe ou les boutons qui démarrent les flux d'informations d'identification tiers sont des contrôles d'interface courants dans les expériences d'authentification. Une expérience d'authentification simple et efficace est une première étape importante pour que les utilisateurs commencent à utiliser votre complément.

## <a name="best-practices"></a>Meilleures pratiques

|À faire|À ne pas faire|
|:----|:----|
|Utilisez l'authentification unique (SSO) pour authentifier les utilisateurs dans votre complément.|Demander aux utilisateurs de se connecter à votre complément séparément de leur compte Microsoft personnel ou de leur compte Office 365 (travail ou école).|
|Avant l'étape de connexion, décrivez la valeur de votre complément ou démontrez sa fonctionnalité sans avoir besoin d'un compte. |Attendez-vous à ce que les utilisateurs se connectent sans comprendre la valeur et les avantages de votre complément.|
|Guidez les utilisateurs à travers l'authentification avec un bouton principal, très visible sur chaque écran. |Attirez l'attention sur les tâches secondaires et tertiaires avec des boutons concurrents et des appels à l'action.|
|Utilisez des noms de bouton décrivant des tâches spécifiques, tels que « Connexion » ou « Créer un compte ».   |Utilisez des noms de bouton vagues tels que « Envoyer » ou « Commencer » pour guider les utilisateurs à travers les flux d'authentification.|
|Utilisez une boîte de dialogue pour attirer l'attention des utilisateurs sur les formulaires d'authentification.    |Surchargez votre volet de tâches avec une première expérience d'exécution et des formulaires d'authentification.|
|Trouvez de petites fonctionnalités dans le flux comme la mise au point automatique sur les boîtes de saisie. |Ajoutez des étapes inutiles à l'interaction, par exemple en demandant aux utilisateurs de cliquer dans les champs de formulaire.|
|Fournir aux utilisateurs un moyen de se déconnecter et de se ré-authentifier.    |Forcer les utilisateurs à désinstaller pour changer d'identité.|

> [!NOTE]
> L’API de l’authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).


## <a name="authentication-flow"></a>Flux d’authentification
Si la connexion unique n'est pas encore disponible pour vos utilisateurs, envisagez un autre flux d'authentification. Donnez aux utilisateurs le choix de se connecter directement à votre service ou à un fournisseur d'identité tel que Microsoft.

1. First Run Placemat - Placez votre bouton de connexion comme une action d'appel claire dans la première expérience d'exécution de votre complément.
![](../images/add-in-fre-value-placemat.png)

2. Boîte de dialogue Choix du fournisseur d'identité - Affiche une liste claire des fournisseurs d'identité, y compris un formulaire pour un nom d'utilisateur et un mot de passe, le cas échéant. Votre interface utilisateur peut être bloquée pendant que la boîte de dialogue d'authentification est ouverte.
![](../images/add-in-auth-choices-dialog.png)



3. Connexion au fournisseur d'identité - Le fournisseur d'identité aura sa propre interface utilisateur. Microsoft Azure Active Directory permet la personnalisation des pages de connexion et d'accès pour une apparence cohérente avec votre service. [En savoir plus](https://docs.microsoft.com/azure/active-directory/fundamentals/customize-branding).
![](../images/add-in-auth-identity-sign-in.png)

4. Progression - Indique la progression lorsque les paramètres et l'interface utilisateur sont chargés.
![](../images/add-in-auth-modal-interstitial.png)

> [!NOTE] 
> Lorsque vous utilisez le service Identité de Microsoft, vous avez la possibilité d'utiliser un bouton de connexion de marque personnalisable selon des thèmes clairs et sombres. En savoir plus.

## <a name="single-sign-on-authentication-flow"></a>Flux d'authentification unique
L'authentification unique est toujours en préversion. Une fois disponible à grande échelle, utilisez-la pour une expérience utilisateur optimale. L'identité de l'utilisateur dans Office est utilisée pour se connecter à votre complément. Par conséquent, les utilisateurs ne se connectent qu'une seule fois. Cela supprime les frictions dans l'expérience, ce qui facilite le démarrage de vos clients.

1. Lors de l'installation d'un complément, un utilisateur verra une fenêtre de consentement semblable à celle ci-dessous : ![](../images/add-in-auth-SSO-consent-dialog.png)
> [!NOTE]
> L'éditeur de complément aura le contrôle sur le logo, les chaînes et les champs d'autorisation inclus dans la fenêtre de consentement. L'interface utilisateur est préconfigurée par Microsoft.

2. Le complément se chargera après que l'utilisateur aura consenti. Il peut extraire et afficher toutes les informations personnalisées nécessaires à l'utilisateur.
![](../images/add-in-ribbon.png)

## <a name="see-also"></a>Voir aussi
- En savoir plus sur [le développement des compléments SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)