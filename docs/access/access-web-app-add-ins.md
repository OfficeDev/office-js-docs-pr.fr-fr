# <a name="create-add-ins-for-access-web-apps"></a>Création de compléments pour les applications web Access

>**Important :** nous ne vous recommandons plus de créer et d’utiliser les bases de données et les applications web Access dans SharePoint. Nous vous recommandons plutôt d’utiliser [Microsoft PowerApps](https://powerapps.microsoft.com/) pour créer des solutions professionnelles sans code pour des appareils mobiles et web.

Cet article explique comment utiliser Visual Studio 2015 pour développer un complément Office qui cible les applications web Access.

>**Remarque :** pour plus d’informations sur le développement de solutions pour Access à l’aide de VBA, consultez la rubrique [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx) sur MSDN.

## <a name="prerequisites"></a>Conditions préalables

Pour créer une Complément Office qui cible applications web Access, vous avez besoin des éléments suivants :

- Visual Studio 2015

- Un site SharePoint Online (inclus dans plusieurs abonnements Office 365). Ce site doit comporter un catalogue de compléments. Pour plus d’informations, reportez-vous à la rubrique [Configuration d’un catalogue de compléments sur SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


>**Remarque :** les compléments Office sont compatibles avec les applications web Access hébergées sur SharePoint Online ou Office 365. L’application de bureau Access 2013 ne prend pas en charge les compléments Office. Les compléments Office qui ciblent les applications Access sont prises en charge par les versions 1.1 et ultérieures d’Office.js.


## <a name="create-a-project-in-visual-studio"></a>Créer un projet dans Visual Studio

1.  Ouvrez Visual Studio et, dans le menu, choisissez **Fichier**, **Nouveau**, **Projet**. La boîte de dialogue **Nouveau projet** s’ouvre.

2. Dans la boîte de dialogue **Nouveau projet**, dans le volet de gauche, accédez à **Installé**, **Modèles**, **Visual C#**, **Office/SharePoint**, **Compléments Office**.

    >**Remarque :**  Si ce modèle n’est pas installé, consultez le billet de blog [Derniers outils de développement Microsoft Office pour Visual Studio 2015](https://blogs.msdn.microsoft.com/visualstudio/2015/11/23/latest-microsoft-office-developer-tools-for-visual-studio-2015/).

3. Dans la boîte de dialogue **Nouveau projet**, dans le volet central, choisissez **Complément Office**.

4. En bas de la boîte de dialogue, entrez un nom pour votre projet et choisissez **OK**. Cette action ouvre la boîte de dialogue **Créer un complément Office**.

5. Dans la boîte de dialogue **Créer un complément Office**, choisissez **Contenu** et cliquez sur **Suivant**.

6. Dans l’écran suivant de la boîte de dialogue  **Créer un complément Office**, choisissez  **Complément de base** ou **Complément de visualisation de document**, et assurez-vous que la case pour  **Access** est cochée.

7. Lorsque vous avez terminé, sélectionnez **Terminer**. Visual Studio créera un projet de démarrage sur lequel baser votre travail.

8. Dans **l’Explorateur de solutions**, choisissez le projet web du projet (**nom_projet > Web**). Dans le volet de propriétés, recherchez l’entrée pour l’**URL SSL**. Cela doit avoir la forme suivante : `https://localhost:44314/`. Sélectionnez cette URL et copiez-la dans le Presse-papiers. Vous en aurez besoin dans peu de temps.

9. Cliquez avec le bouton droit sur le nom de votre projet dans **l’Explorateur de solutions**. Dans le menu contextuel, cliquez sur **Publier**. Cette action ouvre l’assistant **Publier votre complément**.

10. Dans l’assistant **Publier votre complément**, sélectionnez le menu déroulant associé à **Profil actuel**. Dans cette liste déroulante, choisissez **nouveau**. Cette action ouvre la boîte de dialogue **Publier les compléments Office et SharePoint**.

11. Dans cette boîte de dialogue, choisissez **Créer profil**, entrez un nom reconnaissable pour le profil, puis cliquez sur **Terminer**. La boîte de dialogue **Publier des compléments Office et SharePoint** se ferme et vous renvoie à l’assistant **Publier votre complément**.

12. Dans l’assistant, choisissez **Créer un package pour le complément**. Cette action va finaliser votre complément pour le rendre publiable dans un catalogue de compléments SharePoint.

13. Dans la page suivante, pour la rubrique **Où est hébergé votre site web ?**, placez l’URL de l’hôte de votre site web. Il peut s’agir de la valeur **URL SSL** copiée à l’étape 8. Ensuite, cliquez sur **Terminer**.

14. Dans **l’Explorateur de solutions**, cliquez avec le bouton droit de la souris sur le nœud Manifeste du projet (directement sous le nom du projet) et sélectionnez **Ouvrir le dossier dans l’Explorateur de fichiers**. Notez le chemin d’accès à ce fichier. Vous aurez besoin de cette valeur ultérieurement.

>**Remarque :** Pour déboguer le complément, vous devez le déployer avec une application web Access.

## <a name="review-the-manifest-and-the-homehtml-file"></a>Passer en revue le manifeste et le fichier Home.Html

1. Dans votre projet Visual Studio, ouvrez le fichier **Home.html** et trouvez les lignes qui font référence à la bibliothèque de scripts Office.js.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

    >**Remarque :** Cette balise script fait référence à la version 1.1 (et versions ultérieures) d’Office.js. Access utilise les éléments de l’API introduits dans la version 1.1.

2. Ouvrez le fichier manifeste associé à votre projet. Ce fichier sera nommé d’après le nom de votre projet, et aura l’extension « .xml ».

3.  Dans le fichier manifeste, trouvez la section **Hôtes** et recherchez une entrée **Hôte**.

    ```xml
    <Hosts> <Host Name="Database" /> </Hosts>
    ```
    
    >**Remarque :** Il s’agit de l’endroit où sont répertoriées les applications pouvant utiliser le complément. Étant donné que vous avez sélectionné **Access** dans la boîte de dialogue **Créer un complément Office**, l’option **Base de données** est répertoriée. Si vous avez inclus Excel, il existe également une entrée **Classeur**.

Les Compléments Office et SharePoint sont basées sur le web. Le code du complément doit être hébergé sur un serveur web. Pour cet exemple, le serveur web est votre ordinateur de développement. Le serveur doit être en cours d’exécution pour que le complément s’en serve pour les tests, ce qui signifie que Visual Studio doit exécuter le complément au moment où vous le visualisez et le déboguez dans SharePoint.

Pour qu’un utilisateur trouve et utilise le complément, celui-ci doit être inscrit auprès d’un catalogue de compléments dans SharePoint. Les informations dont le catalogue de compléments a besoin sont contenues dans le fichier manifeste.

>**Remarque :**  Vous devrez créer une application web Access pour héberger votre complément Office.

## <a name="publish-your-add-in-to-a-sharepoint-online-catalog"></a>Publier votre complément dans un catalogue SharePoint Online

1.  Connectez-vous à SharePoint Online ou Office 365, puis accédez au **centre d’administration SharePoint** en choisissant **Admin** dans la barre d’outils Office 365 en haut de la page.

2. Sur la page **Centre d’administration SharePoint**, accédez à la barre de liens sur la gauche et cliquez sur **compléments**. Cela permet d’accéder à l’affichage des compléments.

3. Dans le volet central de la page, choisissez **Catalogue des compléments**. Cela permet d’accéder à la page de **Catalogue**.

4. Dans la page  **Catalogue**, choisissez  **Distribuer les compléments Office**. Ceci vous conduit à une page d’annuaire appelée  **Compléments Office** qui répertorie toutes les Compléments Office installées.

5. En haut de la page **Compléments Office**, sélectionnez **nouveau complément**. Cette action permet d’afficher la boîte de dialogue **Ajouter un document**.

6. Dans la boîte de dialogue **Ajouter un document**, cliquez sur **Parcourir**, puis accédez à l’emplacement du fichier manifeste dans votre projet Visual Studio. Si vous avez précédemment copié l’adresse de votre fichier manifeste, vous pouvez la coller dans cette boîte de dialogue.

7. Sélectionnez le fichier manifeste dans votre projet, puis cliquez sur **OK**. SharePoint ajoutera votre complément à la bibliothèque SharePoint locale.

>**Remarque :**  Cette procédure suppose que vous avez créé un site de test sur SharePoint. Dans le cas contraire, vous pouvez le créer à partir de l’onglet **Sites** en haut de la fenêtre de SharePoint. Vous pouvez utiliser une application web Access existante si elle est disponible.

## <a name="create-an-access-web-app-to-host-your-add-in"></a>Créer une application web Access pour héberger votre complément

1. Accédez à votre site de test. Dans la barre de liens de gauche, cliquez sur **Contenu du site**. Cela permet d’accéder à la page **Contenu du site** de votre site de test.

2. Sur la page **Contenu du site**, choisissez la vignette **Ajouter un complément**. Cela vous amène à la page **Contenu du site - Vos compléments**.

3. Sur la page **Contenu du site – Vos compléments**, utilisez la barre de recherche en haut de la page pour rechercher **Application**.

4. Vous devez maintenant voir une mosaïque pour **Application**.

    >**Remarque :**  Gardez à l’esprit qu’il ne s’agit pas de votre complément Office, mais de nouvelles applications web Access. Ces applications web Access vont héberger votre complément Office.

5. Le fait de cliquer sur cette vignette fait apparaître la boîte de dialogue **Ajouter une application Access**. Entrez un nom unique pour votre application Access et cliquez sur **Créer**. SharePoint peut mettre un certain temps à créer votre application. Lorsqu’elle sera terminée, vous verrez votre application Access répertoriée sur la page **Contenu du site**, avec une étiquette **nouveau**.

6. Vous devez ouvrir l’applicationAccess dans la version de bureau de Microsoft Access 2013 et y ajouter des données avant de l’ouvrir et de l’afficher dans SharePoint.

## <a name="add-your-add-in-to-an-access-web-apps"></a>Ajouter votre complément à une applications web Access

1. Ouvrez une application web Access.

2. Dans la barre d’onglets SharePoint, sélectionnez l’icône « engrenage » dans le coin supérieur gauche. Un menu s’affiche. Cliquez sur l’élément de menu **Compléments Office**. Cette action ouvre la boîte de dialogue **Compléments Office**.

3. Choisissez la vue  **Mon organisation** et patientez pendant que SharePoint remplit la boîte de dialogue avec les Compléments Office dont vous disposez.

4. L’un des compléments de la boîte de dialogue doit correspondre au complément Office que vous avez enregistré dans une procédure antérieure. Choisissez ce complément et insérez-le dans vos applications web Access. N’oubliez pas que l’application doit s’exécuter dans Visual Studio pour qu’il soit détecté et qu’il s’affiche sur votre page d’applications web Access.

## <a name="debug-your-add-in-for-office"></a>Déboguer le complément Office

Pour déboguer votre complément, dans Internet Explorer, appuyez sur F12 ou cliquez sur l’icône d’engrenage dans la barre d’onglets des navigateurs (pas l’icône d’engrenage sur la page SharePoint). Ceci affiche les outils de débogage F12 fournis par Internet Explorer 11. Si vous utilisez un autre navigateur, consultez la documentation de votre navigateur pour savoir comment entrer en mode de débogage.

À ce stade, vous pouvez définir des points d’arrêt, parcourir votre code JavaScript, explorer les DOM et modifier le code pour vérifier que vos modifications apparaissent dans le complément Office ciblant les applications web Access. Pour plus d’informations, consultez la rubrique [Utilisation des outils de développement F12](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29).

## <a name="next-steps"></a>Étapes suivantes

Téléchargez l’exemple sur la page [Office 365 : Liaison et manipulation des données dans une application web Access](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e) pour savoir comment implémenter un Complément Office qui manipule des données dans une application web Access.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Présentation de l’API JavaScript pour compléments](../develop/understanding-the-javascript-api-for-office.md)

- [Interface API JavaScript pour Office](http://dev.office.com/reference/add-ins/javascript-api-for-office)
