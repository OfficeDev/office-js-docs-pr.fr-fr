
# <a name="best-practices-for-developing-office-add-ins"></a>Meilleures pratiques en matière de développement de compléments Office


Des compléments efficaces proposent des fonctionnalités uniques et attrayantes qui étendent les applications Office d’une manière visuellement attractive. Pour créer un complément intéressant, offrez une première expérience attractive à vos utilisateurs, concevez une interface utilisateur de premier choix et optimisez les performances de votre complément. Appliquez les meilleures pratiques décrites dans cet article pour créer des compléments permettant aux utilisateurs d’accomplir leurs tâches rapidement et efficacement.

>
  **Remarque :** Lorsque vous créez votre complément, si vous envisagez de le [publier](../publish/publish.md) dans Office Store, assurez-vous que vous respectez les [stratégies de validation Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability)).

## <a name="provide-clear-value"></a>Indication d’une valeur claire

- Créez des compléments qui aident les utilisateurs à réaliser des tâches rapidement et efficacement. Concentrez-vous sur des scénarios adaptés aux applications Office. Par exemple :
 - Réalisez des tâches de création essentielles plus rapidement et plus facilement, avec moins d’interruptions.
 - Développez de nouveaux scénarios dans Office.
 - Intégrez des services complémentaires dans des hôtes Office.
 - Améliorez l’expérience Office pour accroître la productivité.
- Assurez-vous que la valeur de votre complément apparaîtra clairement aux utilisateurs dès la première utilisation en créant une [première expérience enrichissante](#create-an-engaging-first-run-experience).
- Rédigez une [description claire sur l’Office Store](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx). Soulignez les avantages de votre complément dans votre titre et votre description. Ne comptez pas sur votre marque pour communiquer sur les fonctionnalités de votre complément.


## <a name="create-an-engaging-first-run-experience"></a>Création d’une première expérience intéressante



- Attirez de nouveaux utilisateurs avec une première expérience très simple et intuitive. Les utilisateurs décident toujours d’utiliser ou d’abandonner un complément après l’avoir téléchargé à partir du Windows Store.

 - Indiquez clairement les étapes que l’utilisateur doit suivre pour utiliser votre complément. Utilisez des vidéos, des schémas, des panneaux de pagination ou d’autres ressources pour attirer les utilisateurs.

 - N’hésitez pas à ajouter un texte pour insister sur l’utilité de votre complément sur l’écran de connexion des utilisateurs.

 - Proposez une interface utilisateur pédagogique pour guider les utilisateurs et la personnaliser.

    ![Capture d’écran illustrant un complément de volet Office avec des étapes de mise en route en regard d’un complément sans étapes de mise en route](../../images/586202ad-333b-417c-ad31-cc6eb952b239.png)

  - Si votre complément de contenu est lié à des données dans le document de l’utilisateur, incluez des exemples de données ou un modèle pour montrer aux utilisateurs le format de données à utiliser.

    ![Capture d’écran illustrant un complément de contenu avec des données en regard d’un complément de contenu sans données](../../images/7de2215f-ccef-4f82-aa9d-babcbddae0c6.png)

- Offrez des [essais gratuits](https://msdn.microsoft.com/en-us/library/dn456317.aspx#Anchor_1). Si votre complément nécessite un abonnement, proposez certaines fonctionnalités gratuitement.

- Facilitez l’inscription. Préremplissez les informations (e-mail, nom d’affichage) et ignorez les vérifications d’adresses e-mail.

- Évitez d’utiliser des fenêtres contextuelles. Si vous devez les utiliser, aidez les utilisateurs à les activer.

- Utilisez l’[authentification unique (SSO)](../outlook/authenticate-a-user-with-an-identity-token.md).

Pour obtenir les modèles illustrant les modèles de conception à appliquer lors du développement de votre première expérience d’utilisation, voir [Modèles de conception de l’expérience utilisateur pour les compléments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="use-add-in-commands"></a>Utilisation des commandes de complément

- Définissez des accès pertinents à vos compléments dans l’interface utilisateur en utilisant des commandes de complément. Pour plus d’informations, y compris les meilleures pratiques de conception, voir [Commandes de complément](../design/add-in-commands.md).

## <a name="apply-ux-design-principles"></a>Application des principes de conception de l’expérience utilisateur

- Assurez-vous que l’aspect, la convivialité et la fonctionnalité de votre complément améliorent l’expérience Office. Utilisez [Office UI Fabric](https://dev.office.com/fabric).

- Privilégiez le contenu plutôt que l’apparence. Évitez les éléments d’interface utilisateur superflus qui n’ajoutent pas de valeur à l’expérience utilisateur.

- Gardez le contrôle des utilisateurs. Assurez-vous que ces derniers comprennent les décisions importantes et peuvent facilement rétablir des actions effectuées par le complément.

- Utilisez la personnalisation afin d’inspirer la confiance et d’orienter les utilisateurs. N’utilisez pas la personnalisation afin de submerger les utilisateurs ou de faire de la publicité.

- Évitez d’utiliser le défilement. Optimisez votre complément pour une résolution de 1366 x 768.

- N’incluez pas d’image sans licence.

- Utilisez un [langage clair et simple](../design/voice-guidelines.md) dans votre complément.

- Soulignez l’[accessibilité](../design/accessibility-guidelines.md) : votre complément doit être facile à utiliser pour tous les utilisateurs et s’accommoder de technologies d’assistance telles que les lecteurs d’écran.

- Adaptez-le à toutes les plateformes et méthodes d’entrée, y compris la souris/le clavier et la [fonction tactile](#optimize-for-touch). Assurez-vous que votre interface utilisateur réagit à différents formats.

Pour les modèles appliquant des principes de conception que vous pouvez utiliser et personnaliser lors du développement de votre complément, voir [Modèles de conception de l’expérience utilisateur pour les compléments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

### <a name="optimize-for-touch"></a>Optimisation de la fonction tactile



- Utilisez la propriété [Context.touchEnabled](../../reference/shared/office.context.touchenabled.md) pour déterminer si l’application hôte sur laquelle votre complément est exécuté est compatible avec la fonction tactile.

     >**Remarque**  Cette propriété n’est pas prise en charge dans Outlook.
- Assurez-vous que toutes les commandes sont correctement dimensionnées pour l’interaction tactile. Par exemple, vérifiez que les boutons disposent de cibles tactiles adéquates et que les zones de texte sont assez grandes pour permettre la saisie.

- N’utilisez pas de méthodes d’entrée non tactiles comme l’utilisation du curseur ou du clic droit.

- Assurez-vous que votre complément fonctionne dans les modes portrait et paysage. Gardez à l’esprit qu’une partie de votre complément pourrait être masquée par le clavier virtuel sur les appareils tactiles.

- Testez votre complément sur un véritable appareil en utilisant le [chargement de version test](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).


 >**Remarque :** si vous utilisez [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) pour vos éléments de conception, un grand nombre de ces éléments sont pris en charge.


## <a name="optimize-and-monitor-add-in-performance"></a>Optimisation et contrôle des performances du complément



- Donnez l’impression que l’interface utilisateur réagit rapidement. Votre complément doit se charger en 500 ms au maximum.

- Veillez à ce que toutes les interactions utilisateur répondent en moins d’une seconde.

-  Fournissez des indicateurs de chargement pour les opérations à longue durée d’exécution.

- Utilisez un CDN pour héberger les images, les ressources et les bibliothèques communes. Chargez autant d’éléments que possible à partir d’un seul emplacement.

- Suivez les pratiques web standard pour optimiser votre page web. En production, utilisez uniquement les versions réduites des bibliothèques. Chargez uniquement les ressources dont vous avez besoin et optimisez leur chargement.

- Si l’exécution des opérations dure longtemps, fournissez des commentaires aux utilisateurs. Prenez en compte les seuils indiqués dans le tableau suivant. Voir également [Limites des ressources et optimisation des performances pour les compléments Office](../../docs/develop/resource-limits-and-performance-optimization.md).


|**Classe d’interaction**|**Cible**|**Limite supérieure**|**Perception humaine**|
|:-----|:-----|:-----|:-----|
|Instantanée|<= 50 ms|100 ms|Aucun délai notable.|
|Rapide|50-100 ms|200 ms|Délai notable minime. Aucun commentaire n’est nécessaire.|
|Normale|100-300 ms|500 ms|L’opération va assez vite, sans pour autant pouvoir être qualifiée de rapide. Aucun commentaire n’est nécessaire.|
|Réactive|300-500 ms|1 seconde|L’opération n’est pas rapide, mais le système donne l’impression de répondre. Aucun commentaire n’est nécessaire.|
|Continue|> 500 ms|5 secondes|Attente moyenne, le système n’a plus l’air de répondre. Un commentaire peut-être nécessaire.|
|Captive|> 500 ms|10 secondes|Long, mais pas assez pour faire autre chose. Un commentaire peut-être nécessaire.|
|Étendue|> 500 ms|> 10 secondes|Assez long pour faire quelque chose en attendant. Un commentaire peut être nécessaire.|
|Longue durée|> 5 ms|> 1 minute|Les utilisateurs effectueront certainement une autre action.|
- Surveillez l’état de votre service et utilisez la télémétrie pour surveiller le succès d’utilisateur.


## <a name="market-your-add-in"></a>Commercialisation de votre complément



- Publiez votre complément sur l’[Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx) et faites-en la[promotion](http://msdn.microsoft.com/library/b19e21f8-76f5-44e1-9971-bef79cad4c71%28Office.15%29.aspx)sur votre site web. Créez des [listes Office Store efficaces](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

- Utilisez des titres et des descriptifs courts pour le complément. Ils ne doivent pas comporter plus de 128 caractères.

- Rédigez des descriptions brèves et attrayantes pour votre complément. Répondez à la question « Quel problème ce complément résout-il ? ».

- Faites ressortir la proposition de valeur de votre complément dans le titre et la description. Ne comptez pas sur votre marque.

- Créez un site web pour aider les utilisateurs à trouver votre complément et à l’utiliser.

