---
title: Démarrage rapide de l’authentification unique (SSO)
description: Utiliser le générateur Yeoman pour créer un complément Office Node.js qui utilise la connexion unique.
ms.date: 09/07/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: ecbecfd7e475c224451735c7a864f6de2c230d07
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68239377"
---
# <a name="single-sign-on-sso-quick-start"></a>Démarrage rapide de l’authentification unique (SSO)

Dans cet article, vous allez utiliser le générateur Yeoman pour compléments Office pour créer un complément Office pour Excel, Outlook, Word ou PowerPoint qui utilise l’authentification unique (SSO).

> [!NOTE]
> Le modèle d’authentification unique fourni par le générateur Yeoman pour les compléments Office s’exécute uniquement sur localhost et ne peut pas être déployé. Si vous créez un complément Office avec l’authentification unique à des fins de production, suivez les instructions de [La création d’un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))

- La dernière version de[Yeoman](https://github.com/yeoman/yo) et du [Générateur Yeoman Générateur de compléments Office](../develop/yeoman-generator-overview.md). Pour installer ces outils globalement, exécutez la commande suivante via l’invite de commande.

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- Si vous utilisez un Mac et que l'interface de ligne de commande (CLI) Azure n’est pas installée sur votre ordinateur, vous devez installer [Homebrew](https://brew.sh/). Le script de configuration de l’authentification unique exécuté lors de ce démarrage rapide utilise homebrew pour installer l’interface de ligne de commande Azure, puis utilise la CLI pour configurer l’authentification unique dans Azure.

## <a name="create-the-add-in-project"></a>Création du projet de complément

> [!TIP]
> Le générateur Yeoman peut créer un complément Office prenant en charge l’authentification unique pour Excel, Outlook, Word ou PowerPoint avec le type de script JavaScript ou TypeScript. Les instructions suivantes indiquent `JavaScript` et `Excel`, mais vous devez choisir le type de script et l’application client Office les mieux adaptées à votre scénario.

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Sélectionnez un type de projet :** `Office Add-in Task Pane project supporting single sign-on (localhost)`
- **Sélectionnez un type de script :** `JavaScript`
- **Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`
- **Quelle application client Office voulez-vous prendre en charge ?** Choisissez `Excel`, `Outlook`, `Word`ou `Powerpoint`.

:::image type="content" source="../images/yo-office-sso-excel.png" alt-text="Invites et réponses pour le générateur Yeoman dans une interface de ligne de commande.":::

Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Explorer le projet

Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un code pour un complément de volet Office avec authentification unique.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="configure-sso"></a>Configurer l’authentification unique

Maintenant que votre projet de complément est créé et contient le code nécessaire pour faciliter le processus d’authentification unique, effectuez les étapes suivantes pour configurer l’authentification unique pour votre complément.

1. Accédez au dossier racine du projet.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Exécutez la commande suivante pour configurer l’authentification unique pour le complément.

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > Cette commande échouera si votre locataire est configuré pour nécessiter une authentification à deux facteurs. Dans ce scénario, vous devez effectuer manuellement les étapes d’inscription d’application Azure et de configuration de l’authentification unique en suivant toutes les étapes du didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md) .

3. Une fenêtre de navigateur web s’ouvre et vous invite à vous connecter à Azure. Connectez-vous à Azure à l’aide de vos informations d’identification d’administrateur Microsoft 365. Ces informations d’identification sont utilisées pour inscrire une nouvelle application dans Azure et configurer les paramètres requis par l’authentification unique.

    > [!NOTE]
    > Si vous vous connectez à Azure à l’aide d’informations d’identification non-administrateur au cours de cette étape, le script `configure-sso` ne peut pas fournir d’accord d’administrateur pour le complément aux utilisateurs au sein de votre organisation. Par conséquent, l’authentification unique ne sera pas disponible pour les utilisateurs du complément. vous serez invité à vous connecter.

4. Une fois vos informations d'identification saisies, fermez la fenêtre du navigateur et revenez à l'invite de commande. Au fur et à mesure du processus de configuration de l’authentification unique, les messages d’État s’affichent sur la console. Comme décrit dans la section messages de la console, les fichiers du projet de complément que le générateur Yeoman a créé sont automatiquement mis à jour avec les données requises par le processus d’authentification unique.

## <a name="test-your-add-in"></a>Tester votre complément

Si vous avez créé un complément Excel, Word ou PowerPoint, effectuez les étapes décrites dans la section suivante pour l’essayer. Si vous avez créé un complément Outlook, effectuez plutôt les étapes décrites dans la section [Outlook](#outlook) .

### <a name="excel-word-and-powerpoint"></a>Excel, Word et PowerPoint

Effectuez les étapes suivantes pour tester un complément Excel, Word ou PowerPoint.

1. Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et mettez votre complément en sideload dans l’application client Office précédemment sélectionnée.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Quand Excel, Word ou PowerPoint s’ouvre lorsque vous exécutez la commande précédente, vérifiez que vous êtes connecté avec un compte d’utilisateur membre de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso). Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.

3. Dans l’application cliente Office, choisissez l’onglet **Accueil** , puis **sélectionnez Afficher le volet Office** pour ouvrir le volet Office du complément.

    :::image type="content" source="../images/excel-quickstart-addin-3b.png" alt-text="Bouton complément Excel.":::

4. Au bas du volet des tâches, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique.

5. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur client n’a pas donné son consentement pour que le complément accède à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office avec un compte Microsoft valide ou un compte Microsoft 365 Éducation ou Professionnel. Choisissez **Accepter** pour continuer.

    ![Capture d’écran montrant la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence.](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

6. Le complément récupère les informations de profil de l’utilisateur connecté et écrit celui-ci dans le document. L’image suivante montre un exemple d’informations de profil écrites dans une feuille de calcul Excel.

    ![Capture d’écran illustrant les informations de profil utilisateur dans la feuille de calcul Excel.](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a>Outlook

Pour tester un complément Outlook, procédez comme suit.

1. Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet et démarrer le serveur web local.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md) pour charger le complément dans Outlook. Assurez-vous d’être connecté à Outlook avec un utilisateur membre de la même organisation Microsoft 365 que le compte d’administrateur Microsoft 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso). Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.

3. Rédigez un nouveau message dans Outlook.

4. Dans la fenêtre composition des messages, choisissez le bouton **Afficher le volet Office** pour ouvrir le volet Office du complément.

    ![Capture d’écran illustrant la fenêtre Outlook Composer un message et le bouton du ruban du complément mis en évidence.](../images/outlook-sso-ribbon-button.png)

5. Au bas du volet des tâches, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique.

6. Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs. Cela peut se produire lorsque l’administrateur client n’a pas donné son consentement pour que le complément accède à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office avec un compte Microsoft valide ou un compte Microsoft 365 Éducation ou Professionnel. Choisissez **Accepter** pour continuer.

    ![Capture d’écran de la boîte de dialogue des autorisations demandées avec le bouton Accepter mis en évidence.](../images/sso-permissions-request.png)

    > [!NOTE]
    > Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.

7. Le complément récupère les informations du profil de l’utilisateur connecté et les écrit dans le corps de l'e-mail.

    ![Capture d’écran illustrant les informations de profil utilisateur dans la fenêtre Composer un message dans Outlook.](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément de volet Office qui utilise l’authentification unique lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge. Pour en savoir plus sur la personnalisation de votre complément afin d’ajouter une nouvelle fonctionnalité qui requiert des autorisations différentes, voir [Personnaliser votre complément compatible avec l’authentification unique Node.js](sso-quickstart-customize.md).

## <a name="see-also"></a>Voir aussi

- [Activer l’authentification unique pour des compléments Office](../develop/sso-in-office-add-ins.md)
- [Personnaliser votre complément compatible avec l’authentification unique Node.js](sso-quickstart-customize.md)
- [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md)
- [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Utilisation de Visual Studio Code pour publier](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)