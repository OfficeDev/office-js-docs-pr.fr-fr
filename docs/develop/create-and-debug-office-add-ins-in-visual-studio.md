---
title: Créer et déboguer des compléments Office dans Visual Studio
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 3e4fbcd3919be0d5510b36ae77a6e3706eab9689
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437604"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>Créer et déboguer des compléments Office dans Visual Studio

Cet article explique comment utiliser Visual Studio pour créer votre premier complément Office. Les étapes décrites dans cet article concernent Visual Studio 2015. Si vous utilisez une autre version de Visual Studio, les procédures peuvent légèrement varier.

> [!NOTE]
> Si vous débutez avec les compléments pour OneNote, reportez-vous à [Créer votre premier complément OneNote](../onenote/onenote-add-ins-getting-started.md).

## <a name="create-an-office-add-in-project-in-visual-studio"></a>Créer un projet de complément Office dans Visual Studio


Pour commencer, vérifiez que les [outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx) sont installés et que vous disposez d’une version de Microsoft Office. Vous pouvez participer au [programme pour les développeurs Office 365](https://developer.microsoft.com/en-us/office/dev-program), ou suivre ces instructions pour obtenir la [version la plus récente](../develop/install-latest-office-version.md).


1. Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.
    
2. Dans la liste des types de projets sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments web**, puis sélectionnez un des projets de compléments.  
    
3. Nommez le projet, puis cliquez sur **OK** pour créer le projet.
    
4. Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. La page par défaut Home.html s’ouvre dans Visual Studio.
    
Dans Visual Studio 2015, certains des modèles de projet de complément ont été mis à jour pour refléter des fonctionnalités supplémentaires :


- Des compléments de contenu peuvent apparaître dans le corps des documents Access et PowerPoint, en plus des feuilles de calcul Excel. Vous pouvez également choisir l’option de projet de base pour créer un projet de complément ayant un contenu élémentaire avec code de démarrage minimal, ou l’option Projet de visualisation de documents (pour Access et Excel seulement) afin de créer un complément dont le contenu est plus complet, qui inclut un code de démarrage pour visualiser et se lier à des données.
    
- Les compléments Outlook comprennent des options permettant non seulement d’inclure votre complément dans un message électronique ou un rendez-vous, mais aussi d’indiquer si le complément est disponible lorsqu’un message électronique ou un rendez-vous est composé et lu.
    

> [!NOTE]
> Dans Visual Studio, la plupart des options sont compréhensibles par leurs descriptions, sauf la case à cocher **Message électronique**. Cochez cette case si vous souhaitez créer un complément Outlook qui apparaît non seulement avec les éléments de messagerie, mais aussi avec les demandes de réunion, les réponses et les annulations.

Lorsque vous avez terminé l’Assistant, Visual Studio crée une solution qui contient deux projets.



|**Projet**|**Description**|
|:-----|:-----|
|Projet de complément|Contient seulement un fichier de manifeste XML, qui contient tous les paramètres qui décrivent votre complément. Ces paramètres aident l’hôte Office à déterminer quand votre complément doit être activé et où il doit apparaître. Visual Studio génère le contenu de ce fichier pour vous afin que vous puissiez exécuter le projet et utiliser immédiatement votre complément. Vous pouvez modifier ces paramètres à tout moment à l’aide de l’éditeur de manifeste.|
|Projet d’application web|Contient les pages de contenu de votre complément, notamment tous les fichiers et références de fichiers dont vous avez besoin pour développer des pages HTML et JavaScript compatibles avec Office. Pendant que vous développez votre complément, Visual Studio héberge l’application web sur votre serveur IIS local. Lorsque vous êtes prêt à la publier, vous devez trouver un serveur pour héberger ce projet.Pour en savoir plus sur les projets d’applications web ASP.NET, voir [Projets web ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## <a name="modify-your-add-in-settings"></a>Modifier les paramètres de votre complément


Pour modifier les paramètres de votre complément, modifiez le fichier manifeste XML du projet. Dans l’**Explorateur de solutions**, développez le nœud de projet du complément et le dossier contenant le manifeste XML, puis sélectionnez le manifeste XML. Vous pouvez pointer sur n’importe quel élément du fichier pour afficher une info-bulle qui décrit l’objectif de l’élément. Pour plus d’informations sur le fichier manifeste, voir l’article sur le [manifeste XML de compléments Office](../develop/add-in-manifests.md).


## <a name="develop-the-contents-of-your-add-in"></a>Développer le contenu de votre complément


Alors que le projet de complément vous permet de modifier les paramètres qui décrivent le complément, l’application web fournit le contenu qui apparaît dans le complément. 

Le projet d’application web contient une page HTML par défaut et le fichier JavaScript que vous pouvez utiliser pour commencer. Il contient également un fichier JavaScript commun à toutes les pages que vous ajoutez à votre projet. Ces fichiers sont pratiques car ils contiennent des références à d’autres bibliothèques JavaScript, notamment l’API JavaScript pour Office. 

Au fur et à mesure que votre complément devient plus complexe, vous pouvez ajouter d’autres fichiers HTML et JavaScript. Vous pouvez utiliser le contenu des fichiers HTML et JavaScript par défaut comme exemples des types de références que vous pouvez ajouter à d’autres pages de votre projet pour les faire fonctionner avec votre complément. Le tableau suivant décrit les fichiers HTML et JavaScript par défaut.



|**Fichier**|**Description**|
|:-----|:-----|
|**Home.html**|Situé dans le dossier  **de base** du projet ; il s’agit de la page HTML par défaut du complément. Cette page apparaît en tant que première page du complément lorsqu’elle est activée dans un élément de rendez-vous, de message électronique ou de document. Ce fichier est utile, car il contient toutes les références de fichiers dont vous avez besoin pour commencer. Lorsque vous êtes prêt à créer votre premier complément, ajoutez votre code HTML à ce fichier.|
|**Home.js**|Situé dans le dossier  **de base** du projet ; il s’agit du fichier JavaScript associé à la page Home.js. Vous pouvez placer tout code propre au comportement de la page Home.html dans le fichier Home.js. Ce dernier contient un exemple de code pour vous aider à commencer.|
|**App.js**|Situé dans le dossier  **Complément** du projet ; il s’agit du fichier JavaScript par défaut de l’ensemble du complément. Vous pouvez placer tout code commun au comportement de plusieurs pages de votre application dans le fichier App.js. Ce dernier contient un exemple de code pour vous aider à commencer.|

> [!NOTE]
> Vous n’êtes pas obligé d’utiliser ces fichiers. N’hésitez pas à ajouter d’autres fichiers au projet et à les utiliser à la place. Si vous souhaitez voir apparaître un autre fichier HTML comme page initiale du complément, ouvrez l’éditeur de manifeste et définissez la propriété **SourceLocation** sur le nom du fichier.


## <a name="debug-your-add-in"></a>Déboguer votre complément


Lorsque vous êtes prêt à démarrer votre complément, vérifiez les propriétés liées à la génération et au débogage, puis démarrez la solution.


### <a name="review-the-build-and-debug-properties"></a>Réviser les propriétés de génération et de débogage

Avant de démarrer la solution, assurez-vous que Visual Studio va ouvrir l’application hôte souhaitée. Cette information apparaît dans les pages de propriétés du projet avec d’autres propriétés liées à la génération et au débogage du complément.


### <a name="to-open-the-property-pages-of-a-project"></a>Pour ouvrir les pages de propriétés d’un projet


1. Dans l’ **Explorateur de solutions**, choisissez le nom du projet.
    
2. Dans la barre de menus, choisissez  **Affichage**,  **Fenêtre Propriétés**.
    
Le tableau suivant décrit les propriétés du projet.



|**Propriété**|**Description**|
|:-----|:-----|
|**Action de démarrage**|Indique si votre complément doit être débogué dans un client de bureau Office ou dans un client Office Online dans le navigateur spécifié.|
|**Document de démarrage** (compléments de contenu et du volet Office uniquement)|Spécifie le document à ouvrir lors du démarrage du projet.|
|**Projet Web**|Spécifie le nom du projet web associé au complément.|
|**Adresse de messagerie** (compléments Outlook uniquement)|Spécifie l’adresse de messagerie du compte d’utilisateur dans Exchange Server ou Exchange Online avec lequel vous souhaitez tester votre complément Outlook.|
|**URL EWS** (compléments Outlook uniquement)|URL de service web Exchange (par exemple : https://www.contoso.com/ews/exchange.aspx). |
|**URL OWA** (compléments Outlook uniquement)|URL d’application web Outlook (par exemple : https://www.contoso.com/owa).|
|**Nom d’utilisateur** (compléments Outlook uniquement)|Spécifie le nom de votre compte d’utilisateur dans Exchange Server ou Exchange Online.|
|**Fichier du projet**|Indique le nom du fichier contenant la version, la configuration et d’autres informations sur le projet.|
|**Dossier du projet**|Emplacement du fichier de projet.|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>Utiliser un document existant pour déboguer le complément (compléments de contenu et du volet Office uniquement)


Vous pouvez ajouter des documents au projet de complément. Si vous disposez d’un document qui contient des données de test que vous souhaitez utiliser avec votre application, Visual Studio ouvre ce document lorsque vous commencez le projet.


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>Pour utiliser un document existant pour déboguer le complément


1. Dans l’ **Explorateur de solutions**, choisissez le dossier du projet de complément.
    
    > [!NOTE]
    > Choisissez le projet de complément et non le projet d’application web.

2. Dans le menu **Projet**, choisissez **Ajouter un élément existant**.
    
3. Dans la boîte de dialogue  **Ajouter un élément existant**, recherchez et sélectionnez le document que vous souhaitez ajouter.
    
4. Choisissez le bouton  **Ajouter** pour ajouter le document à votre projet.
    
5. Dans l’ **Explorateur de solutions**, ouvrez le menu contextuel du projet, puis choisissez  **Propriétés**.
    
    Les pages des propriétés relatives au projet s’affichent.
    
6. Dans la liste  **Document de démarrage**, choisissez le document que vous avez ajouté au projet, puis cliquez sur le bouton  **OK** pour fermer les pages de propriétés.
    

### <a name="start-the-solution"></a>Démarrer la solution


Visual Studio génère automatiquement la solution lorsque vous la démarrez. Vous pouvez la démarrer à partir de la barre de  **Menu** en choisissant **Débogage**,  **Démarrer**. 


> [!NOTE]
> Si le débogage de script n’est pas activé dans Internet Explorer, vous ne pourrez pas démarrer le débogueur dans Visual Studio. Pour activer le débogage de script, ouvrez la boîte de dialogue **Options Internet**, choisissez l’onglet **Avancé**, puis désélectionnez les cases à cocher **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.

Visual Studio génère le projet et effectue les actions suivantes.


1. Il crée une copie du fichier manifeste XML et l’ajoute au répertoire  _ProjectName_\Output. L’application hôte utilise cette copie lorsque vous démarrez Visual Studio et déboguez l’application.
    
2. Il crée un ensemble d’entrées dans le Registre de votre ordinateur qui permettent au complément d’apparaître dans l’application hôte.
    
3. Il génère le projet d’application web, puis le déploie sur le serveur web IIS local (http://localhost). 
    
Visual Studio effectue ensuite les actions suivantes :


1. Il modifie l'élément [emplacement source](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) du fichier manifeste XML en remplaçant le jeton ~remoteAppUrl par l'adresse complète de la page de démarrage (par exemple, http://localhost/MyAgave.html).
    
2. Il démarre le projet d’application web dans IIS Express.
    
3. Il ouvre l’application hôte. 
    
Visual Studio n’affiche pas les erreurs de validation dans la fenêtre  **OUTPUT** lorsque vous générez le projet. Visual Studio signale au fur et à mesure les erreurs et les avertissements dans la fenêtre **ERRORLIST**. Visual Studio signale également les erreurs de validation en affichant des traits de soulignement ondulés (appelés aussi zigzags) de différentes couleurs dans le code et l’éditeur de texte. Ces marques sont là pour vous indiquer les problèmes détectés par Visual Studio dans votre code. Pour plus d’informations, voir la page relative au [code et à l’éditeur de texte](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). Pour plus d’informations sur l’activation ou la désactivation de la validation, voir les rubriques suivantes : 

- [Options, Éditeur de texte, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)
    
- [Procédure : définir des options de validation pour l’édition HTML dans Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)
    
- [Validation, CSS, Éditeur de texte, boîte de dialogue Options](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)
    
Pour réviser les règles de validation du fichier manifeste XML dans votre projet, voir [Manifeste XML des compléments Office](../develop/add-in-manifests.md).


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a>Afficher un complément dans Excel, Word ou Project, et avancer pas à pas dans votre code


Si vous définissez la propriété  **Document de démarrage** du projet de complément sur Excel ou Word, Visual Studio crée un document et le complément apparaît. Si vous définissez la propriété **Document de démarrage** du projet de complément afin d’utiliser un document existant, Visual Studio ouvre le document, mais vous devez insérer manuellement le complément. Si vous définissez la propriété **Document de démarrage** sur **Microsoft Project**, vous devez également insérer le complément manuellement.


### <a name="to-show-an-office-add-in-in-excel-or-word"></a>Pour afficher une Complément Office dans Excel ou Word


1. Dans Excel ou Word, dans l’onglet  **Insertion**, choisissez  **Compléments Office**.
    
2. Dans la liste qui apparaît, choisissez votre complément.
    

### <a name="to-show-an-office-add-in-in-project"></a>Pour afficher une Complément Office dans Project


1. Dans Project, dans l’onglet  **Projet**, choisissez  **Compléments Office**.
    
2. Dans la liste qui apparaît, choisissez votre complément.
    
Dans Visual Studio, vous pouvez définir des points d’interruption, puis pendant l’interaction avec votre complément et l’exécution pas à pas du code de vos fichiers HTML, JavaScript et C# ou VB.


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>Afficher le complément Outlook dans Outlook et avancer pas à pas dans votre code


Pour voir le complément dans Outlook, ouvrez un message électronique ou un élément de rendez-vous.

Outlook active le complément pour l’élément à condition que les critères d’activation soient respectés. La barre complément apparaît en haut de la fenêtre de l’inspecteur ou du volet de lecture, et votre complément Outlook apparaît sous la forme d’un bouton dans la barre du complément. Si votre complément est doté d’une commande, un bouton apparaît dans le ruban (soit dans l’onglet par défaut, soit dans un onglet personnalisé indiqué), et le complément n’apparaît pas dans la barre complément.

Pour voir votre complément Outlook, cliquez sur le bouton correspondant.

Dans Visual Studio, vous pouvez définir des points d’interruption, puis pendant l’interaction avec votre complément Outlook et l’exécution pas à pas du code de vos fichiers HTML, JavaScript et C# ou VB. 

Vous pouvez également modifier votre code et vérifier les effets de ces modifications dans votre complément Outlook sans devoir fermer le Complément Office ni redémarrer le projet. Dans Outlook, ouvrez simplement le menu contextuel du complément Outlook, puis choisissez **Recharger**.


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>Modifier le code et continuer le débogage du complément sans redémarrer le projet


Vous pouvez modifier votre code et vérifier les effets de ces modifications dans votre complément sans avoir à fermer l’application hôte et à redémarrer le projet. Après avoir modifié votre code, ouvrez le menu contextuel du complément, puis choisissez  **Recharger**. Quand vous rechargez le complément, il est déconnecté du débogueur Visual Studio. Vous pouvez constater les effets de vos modifications, mais vous ne pouvez pas parcourir pas à pas le code tant que vous n’attachez pas le débogueur Visual Studio à tous les processus Iexplore.exe disponibles.


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a>Pour attacher le débogueur Visual Studio à tous les processus Iexplore.exe disponibles


1. Dans Visual Studio, choisissez  **DÉBOGUER**,  **Attacher au processus**.
    
2. Dans la boîte de dialogue  **Attacher au processus**, choisissez tous les processus  **Iexplore.exe** disponibles, puis sélectionnez le bouton **Attacher**.
    

## <a name="next-steps"></a>Étapes suivantes

- [Déploiement et publication de votre complément Office](../publish/publish.md)
    
