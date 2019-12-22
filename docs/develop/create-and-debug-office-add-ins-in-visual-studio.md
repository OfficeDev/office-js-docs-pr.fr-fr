---
title: Créer et déboguer des compléments Office dans Visual Studio
description: Utiliser Visual Studio pour créer et déboguer des compléments Office dans le client de bureau Office sous Windows
ms.date: 12/16/2019
localization_priority: Priority
ms.openlocfilehash: 2a32075420355e1b70c91c676baf00bc202b18b1
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814123"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Créer et déboguer des compléments Office dans Visual Studio

Cet article décrit comment utiliser Visual Studio 2019 pour créer un complément Office pour Excel, Word, PowerPoint ou Outlook et déboguer le complément dans le client de bureau Office sur Windows. Si vous utilisez une autre version de Visual Studio, les procédures peuvent légèrement varier.

> [!NOTE]
> Visual Studio ne prend pas en charge la création de compléments Office pour OneNote ou un projet, mais vous pouvez utiliser le [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office) pour créer ce genre de compléments.
> - Si vous débutez avec les compléments pour OneNote, reportez-vous à [Créer votre premier complément OneNote](../quickstarts/onenote-quickstart.md).
>
> - Pour commencer à utiliser un complément pour Project, voir [Créer votre premier complément Project](../quickstarts/project-quickstart.md).

## <a name="prerequisites"></a>Conditions préalables

- [Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée

    > [!TIP]
    > Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée. Si cette charge de travail n’est pas encore installée, utilisez Visual Studio Installer pour l’[installer](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).

- Office 2013 ou version ultérieure

    > [!TIP]
    > Si vous n’avez pas Office, vous pouvez rejoindre le[programme Office 365 pour les développeurs](https://developer.microsoft.com/office/dev-program) pour obtenir un abonnement Office 365, ou vous pouvez[vous inscrire à un essai gratuit de 1 mois](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

## <a name="create-the-add-in-project-in-visual-studio"></a>Créer un projet de complément Office dans Visual Studio

Commencer en complétant ces trois étapes, puis suivez les étapes de la section suivante qui correspond au type de complément que vous créez. 

1. Ouvrez Visual Studio, puis dans la barre de menus Visual Studio, choisissez **Créer un nouveau projet**.

2. À l’aide de la zone de recherche, entrez **Compléments**, puis choisissez le type de projet de complément que vous souhaitez créer.

3. Nommez le projet, puis cliquez sur **OK**.

### <a name="word-web-add-in-or-outlook-web-add-in"></a>Complément web Word ou complément web Outlook

Si vous avez choisi de créer un **complément web Word** ou un **complément web Outlook**, Visual Studio crée une solution et ses deux projets s’affichent dans **l’Explorateur de solutions**. Ensuite, vous pouvez [explorer la solution Visual Studio](#explore-the-visual-studio-solution).

### <a name="powerpoint-web-add-in"></a>Complément web PowerPoint

Si vous avez choisi de créer un **complément web PowerPoint**, la boîte de dialogue**créer un complément Office** s’affiche.

- Pour créer un complément de volet tâche, sélectionnez **ajouter de nouvelles fonctionnalités à PowerPoint**, puis cliquez sur le bouton**Terminer** pour créer la solution Visual Studio.

- Pour créer un complément de contenu, sélectionnez **ajouter du contenu à des diapositives PowerPoint**, puis cliquez sur le bouton**Terminer** pour créer la solution Visual Studio.

Ensuite, vous pouvez [explorer la solution Visual Studio](#explore-the-visual-studio-solution).

### <a name="excel-web-add-in"></a>Complément web Excel 

Si vous avez choisi de créer un **complément web Excel**, la boîte de dialogue**créer un complément Office** s’affiche. 

- Pour créer un complément de volet tâche, sélectionnez **ajouter de nouvelles fonctionnalités à Excel**, puis cliquez sur le bouton**Terminer** pour créer la solution Visual Studio.

- Pour créer un complément de contenu, sélectionnez **ajouter du contenu à des tableaux Excel**, puis cliquez sur le bouton**Suivant**, sélectionnez une des options suivantes puis cliquez sur le bouton**Terminer** pour créer la solution Visual Studio :

    - **Complément de base** : pour créer un projet complément de contenu avec un code de démarrage minimal

    - **Complément de visualisation document** : pour créer un projet complément de contenu avec un code starter pour visualiser et lier aux données  

### <a name="explore-the-visual-studio-solution"></a>Explorer la solution Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## <a name="modify-your-add-in-settings"></a>Modifier les paramètres de votre complément

Pour modifier les paramètres de votre complément, modifiez le fichier manifeste XML du projet de complément. Dans l’**Explorateur de solutions**, développez le nœud de projet du complément et le dossier contenant le manifeste XML, puis sélectionnez le manifeste XML. Vous pouvez pointer sur n’importe quel élément du fichier pour afficher une info-bulle qui décrit l’objectif de l’élément. Pour plus d’informations sur le fichier manifeste, voir l’article sur le [manifeste XML de compléments Office](../develop/add-in-manifests.md).

## <a name="develop-the-contents-of-your-add-in"></a>Développer le contenu de votre complément

Alors que le projet de complément vous permet de modifier les paramètres qui décrivent le complément, l’application web fournit le contenu qui apparaît dans le complément. 

Le projet d’application web contient un fichier HTML, un fichier JavaScript et un fichier CSS par défaut que vous pouvez utiliser pour commencer. Certains de ces fichiers contiennent des références à d’autres bibliothèques JavaScript de l’API JavaScript pour Office. Vous pouvez développer votre complément en mettant à jour ces fichiers et/ou en ajoutant des fichiers HTML et JavaScript plus. Le tableau suivant décrit les fichiers par défaut qui sont contenus dans le projet d’application web lors de la création de la solution Visual Studio.

|**Nom de fichier**|**Description**|
|:-----|:-----|
|**Home.html**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.html**<br/>(Outlook)|La page HTML par défaut du complément. Cette page s’affiche comme première page dans le complément, lorsqu’elle est activée dans un document, un message électronique ou un élément de rendez-vous. Ce fichier contient toutes les références de fichier dont vous avez besoin pour commencer. Vous pouvez commencer à développer votre complément en ajoutant votre code HTML dans ce fichier.|
|**Home.js**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.js**<br/>(Outlook)|Fichier JavaScript associé à la page**Home.html** (Excel, PowerPoint, Word) ou la page **MessageRead.html**(Outlook). Ce fichier doit contenir tout code qui est propre au comportement de la page**Home.html** (Excel, PowerPoint, Word) ou de la page **MessageRead.html** (Outlook). Ce fichier contient des exemples de code pour vous aider à démarrer.|
|**Home.CSS**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.css**<br/>(Outlook)|Définit les styles par défaut à appliquer à votre complément. Nous vous recommandons d’utiliser la structure de l’interface utilisateur Office pour la conception et le style. Pour plus d’informations, voir [Structure d’interface utilisateur Office pour compléments Office](../design/office-ui-fabric.md).|

> [!NOTE]
> Vous n’êtes pas obligé d’utiliser ces fichiers. N’hésitez pas à ajouter d’autres fichiers au projet et les utiliser à la place. Si vous souhaitez qu’un autre fichier HTML apparaisse comme page initiale du complément, ouvrez l’éditeur de manifeste et définissez la propriété**SourceLocation** sur le nom du fichier.

## <a name="debug-your-add-in"></a>Déboguer votre complément

Vous pouvez utiliser Visual Studio pour déboguer votre complément dans le client de bureau Office sur Windows, comme décrit dans les sections suivantes :

- [Activer le débogage pour les commandes de compléments et les codes sans interface utilisateur](#enable-debugging-for-add-in-commands-and-ui-less-code)
- [Passez en revue les propriétés de création et débogage](#review-the-build-and-debug-properties)
- [Utiliser un document existant pour déboguer le complément](#use-an-existing-document-to-debug-the-add-in)
- [Démarrer le projet](#start-the-project)
- [Déboguer le code d’un complément Excel, PowerPoint ou Word](#debug-the-code-for-an-excel-powerpoint-or-word-add-in)
- [Déboguer le code d’un complément Outlook](#debug-the-code-for-an-outlook-add-in)

> [!NOTE]
> Vous ne pouvez pas utiliser Visual Studio pour déboguer des compléments Office dans Office sur le web ou Mac. Pour plus d’informations sur le débogage sur ces plateformes, voir [Déboguer les compléments Office dans Office sur le web](../testing/debug-add-ins-in-office-online.md) ou [Déboguer les compléments Office sur iPad et Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

### <a name="enable-debugging-for-add-in-commands-and-ui-less-code"></a>Activer le débogage pour les commandes de compléments et les codes sans interface utilisateur

Lors du débogage d’Office sous Windows par Visual Studio, le complément est hébergé dans une instance du navigateur Microsoft Internet Explorer ou Microsoft Edge. Pour identifier le navigateur utilisé sur votre ordinateur de développement, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### <a name="review-the-build-and-debug-properties"></a>Réviser les propriétés de création et de débogage

Avant de commencer le débogage, passez en revue les propriétés de chaque projet afin de confirmer que Visual Studio ouvre l’application hôte souhaitée et que les autres propriétés de création et débogage sont définies de façon appropriée.

#### <a name="add-in-project-properties"></a>Propriétés du projet de complément

Ouvrir la fenêtre **Propriétés** pour le projet complément pour examiner les propriétés du projet :

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

2. Dans la barre de menu, choisissez **Affichage** >  **Fenêtre Propriétés**.

Le tableau suivant décrit les propriétés du projet de complément.

|**Propriété**|**Description**|
|:-----|:-----|
|**Action de démarrage**|Spécifie le mode de débogage pour votre complément. Actuellement, seul le mode **Client Office pour bureau** est pris en charge pour les projets complément Office.|
|**Document de démarrage**<br/>(Compléments Excel, PowerPoint et Word uniquement)|Spécifie le document à ouvrir lors du démarrage du projet.|
|**Projet Web**|Spécifie le nom du projet web associé au complément.|
|**Adresse e-mail**<br/>(Compléments Outlook uniquement)|Spécifie l’adresse de messagerie du compte utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.|
|**Url EWS**<br/>(Compléments Outlook uniquement)|URL de service web Exchange (par exemple :`https://www.contoso.com/ews/exchange.aspx`). |
|**Url OWA**<br/>(Compléments Outlook uniquement)|URL Outlook sur le web (par exemple : `https://www.contoso.com/owa`).|
|**Utiliser l’authentification multi-facteur**<br/>(Compléments Outlook uniquement)|Valeur booléenne qui indique si l’authentification multi-facteur doit être utilisée.|
|**Nom d'utilisateur**<br/>(Compléments Outlook uniquement)|Spécifie le nom du compte utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Emplacement du fichier de projet.|

> [!NOTE]
> Pour un complément Outlook, vous pouvez choisir de spécifier des valeurs pour une ou plusieurs des propriétés du*complément Outlook uniquement*dans la fenêtre**propriétés**, mais cette opération n’est pas obligatoire.

#### <a name="web-application-project-properties"></a>Propriétés du projet application Web

Ouvrir la fenêtre**Propriétés** pour le projet complément web pour examiner les propriétés du projet :

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément web.

2. Dans la barre de menu, choisissez **Affichage** >  **Fenêtre Propriétés**.

Le tableau suivant décrit les propriétés du projet d’application web qui sont les plus pertinentes aux projets complément Office.

|**Propriété**|**Description**|
|:-----|:-----|
|**SSL activé**|Spécifie si SSL est activé sur le site. Cette propriété doit être définie sur **vrai** pour les projets complément Office.|
|**URL SSL**|Spécifie l’URL HTTPS sécurité pour le site. Lecture seule.|
|**URL**|Spécifie l’URL HTTP pour le site. Lecture seule.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Précise l’emplacement du fichier de projet. Lecture seule. Le fichier manifeste créé par Visual Studio lors de l’exécution est écrit le `bin\Debug\OfficeAppManifests` dossier dans cet emplacement.|

### <a name="use-an-existing-document-to-debug-the-add-in"></a>Utiliser un document existant pour déboguer le complément

Si vous avez un document qui contient les données de test à utiliser pendant le débogage de votre complément Excel, PowerPoint ou Word, Visual Studio peut être configuré pour ouvrir ce dernier lorsque vous démarrez le projet. Pour spécifier un document existant à utiliser pour déboguer le complément, procédez comme suit.

1. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

2. Dans la barre de menus, sélectionnez **Project** > **ajouter un élément existant**.

3. Dans la boîte de dialogue **Ajouter un élément existant**, recherchez et sélectionnez le document que vous souhaitez ajouter.

4. Choisissez le bouton**Ajouter** pour ajouter le document à votre projet.

5. Dans **l’Explorateur de solutions**, choisissez le projet de complément (*pas* le projet d’application web).

6. Dans la barre de menu, choisissez **Affichage** > **Fenêtre Propriétés**.

7. Dans la fenêtre**propriétés**, choisissez la liste**Document de démarrage** et sélectionnez le document que vous avez ajouté au projet. Le projet est désormais configuré pour démarrer le complément dans ce document.

### <a name="start-the-project"></a>Démarrer le projet

Démarrez le projet en choisissant **déboguer** > **démarrer le débogage** à partir de la barre de menus. Visual Studio créera automatiquement la solution et démarrera Office pour héberger votre complément.

> [!NOTE]
> Lorsque vous commencez un projet de complément Outlook, vous serez invité à indiquer vos informations de connexion. Si vous êtes invité à vous connecter à plusieurs reprises ou si vous recevez un message d’erreur indiquant que vous n’êtes pas autorisé, il se peut que l’authentification de base soit désactivée pour les comptes sur votre client Office 365. Dans ce cas, essayez d’utiliser un compte Microsoft à la place. Il se peut également que vous deviez définir la propriété « Utiliser l’authentification multifacteur » sur Vrai dans la boîte de dialogue Propriétés du complément Outlook Web.

Visual Studio génère le projet et effectue les actions suivantes :

1. Crée une copie du fichier manifeste XML et ajoute celui-ci au `_ProjectName_\bin\Debug\OfficeAppManifests` répertoire. L’application hôte consomme cette copie lorsque vous démarrez Visual Studio et débogue le complément.

2. Crée un ensemble d’entrées dans le registre de votre ordinateur qui permettent au complément d’apparaître dans l’application hôte.

3. Génère le projet d’application web, puis le déploie sur le serveur web IIS local (https://localhost).

4. S’il s’agit du premier projet de complément que vous déployez sur un serveur web IIS local, il se peut que vous soyez invité à installer un certificat auto-signé pour le magasin de certificats racines de confiance de l’utilisateur actuel. Cela est nécessaire pour qu’IIS Express puisse afficher correctement le contenu de votre complément.


> [!NOTE]
> La dernière version d’Office peut utiliser un contrôle web plus récent pour afficher le contenu du complément lors de l’exécution de celui-ci sur Windows 10. Si tel est le cas, Visual Studio peut vous inviter à ajouter une exemption de bouclage de réseau local. Cela est nécessaire pour que le contrôle web dans l’application hôte Office puisse accéder au site web déployé sur le serveur web IIS local. Vous pouvez également modifier ce paramètre à tout moment dans Visual Studio sous **Outils** > **Options** > **Outils Office (web)** > **Débogage de compléments web**.


Visual Studio effectue ensuite les actions suivantes :

1. Il modifie l’élément [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) du fichier manifeste XML en remplaçant le jeton`~remoteAppUrl`par l’adresse complète de la page de démarrage (par exemple`https://localhost:44302/Home.html`).

2. Il démarre le projet d’application web dans IIS Express.

3. Il ouvre l’application hôte.

Visual Studio n’affiche pas les erreurs de validation dans la fenêtre **Output** lorsque vous créez le projet. Visual Studio signale les erreurs et avertissements dans la fenêtre **ERRORLIST** lorsqu’elles se produisent. Visual Studio signale également des erreurs de validation en affichant les soulignements ondulés de différentes couleurs (également connus sous soulignements ondulés) dans l’éditeur de code et de texte. Ces marques signalent l’arrivée de problèmes Visual Studio détectés dans votre code. Pour plus d’informations sur comment activer ou désactiver la validation, voir [Options, éditeur de texte, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019).

Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md).

### <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>Déboguer le code d’un complément Excel, PowerPoint ou Word

Si votre complément n’apparaît pas dans le document qui s’affiche dans l’application hôte (Excel, PowerPoint ou Word) après avoir [démarré le projet](#start-the-project), lancez manuellement le complément dans l’application hôte. Par exemple, démarrez votre complément volet tâche en choisissant le bouton**Afficher le volet de tâches** dans l’onglet **Accueil**. Une fois que votre complément est affiché dans Excel, PowerPoint ou Word, vous pouvez déboguer votre code en procédant comme suit :

1. Dans Excel, PowerPoint ou Word, sélectionnez l’onglet **insérer**, puis cliquez sur la flèche vers le bas située à droite de **Mes compléments**.

    ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)

2. Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez votre complément pour effectuer cette opération.

3. Dans Visual Studio, définissez des points d’arrêt dans votre code.

4. Dans Excel, PowerPoint ou Word, interagissez avec votre complément.

5. Lorsque des points d’arrêt sont marqués dans Visual Studio, parcourez le code si besoin.

Vous pouvez modifier votre code et passer en revue les effets de ces modifications dans votre complément sans avoir à fermer l’application hôte et redémarrer le projet. Une fois que vous enregistrez des modifications à votre code, rechargez simplement le complément dans l’application hôte. Par exemple, rechargez un complément de volet de tâches en choisissant le coin supérieur droit du volet Office pour activer la [menu personnalisé](../design/task-pane-add-ins.md#personality-menu), puis**Recharger**.

### <a name="debug-the-code-for-an-outlook-add-in"></a>Déboguer le code d’un complément Outlook

Une fois que vous avez [démarré le projet](#start-the-project) et que Visual Studio lance Outlook pour héberger votre complément, ouvrez un élément de courrier électronique ou un rendez-vous. 

Outlook active le complément pour l’élément à condition que les critères d’activation soient respectés. La barre complément apparaît en haut de la fenêtre de l’inspecteur ou du volet de lecture, et votre complément Outlook apparaît sous la forme d’un bouton dans la barre du complément. Si votre complément est doté d’une commande, un bouton apparaît dans le ruban (soit dans l’onglet par défaut, soit dans un onglet personnalisé indiqué), et le complément n’apparaît pas dans la barre complément.

Pour voir votre complément Outlook, cliquez sur le bouton correspondant. Une fois que votre complément est affiché dans Outlook, vous pouvez déboguer votre code en procédant comme suit :

1. Dans Visual Studio, définissez des points d’arrêt dans votre code.

2. Dans Outlook, interagissez avec votre complément.

3. Lorsque des points d’arrêt sont marqués dans Visual Studio, parcourez le code si besoin.

Vous pouvez modifier votre code et passer en revue les effets de ces modifications dans votre complément sans avoir à fermer Outlook et redémarrer le projet. Une fois que vous enregistrez des modifications à votre code, il vous suffit d’ouvrir le menu contextuel pour le complément (dans Outlook), puis **recharger**.

## <a name="next-steps"></a>Étapes suivantes

Une fois que votre complément fonctionne comme vous le souhaitez, voir [Déployer et publier votre complément Office](../publish/publish.md) pour en savoir plus sur les méthodes avec lesquelles vous pouvez distribuer le complément aux utilisateurs.
